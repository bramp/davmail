/*
 * DavMail POP/IMAP/SMTP/CalDav/LDAP Exchange Gateway
 * Copyright (C) 2009  Mickael Guessant
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
 */
package davmail.exchange;

import davmail.BundleMessage;
import davmail.Settings;
import davmail.exception.DavMailAuthenticationException;
import davmail.exception.DavMailException;
import davmail.exception.WebdavNotAvailableException;
import davmail.exchange.condition.Condition;
import davmail.exchange.condition.MultiCondition;
import davmail.exchange.entity.*;
import davmail.http.DavGatewayHttpClientFacade;
import davmail.http.DavGatewayOTPPrompt;
import davmail.util.StringUtil;
import org.apache.commons.httpclient.*;
import org.apache.commons.httpclient.methods.GetMethod;
import org.apache.commons.httpclient.methods.PostMethod;
import org.apache.commons.httpclient.params.HttpClientParams;
import org.apache.log4j.Logger;
import org.htmlcleaner.CommentNode;
import org.htmlcleaner.ContentNode;
import org.htmlcleaner.HtmlCleaner;
import org.htmlcleaner.TagNode;

import javax.imageio.ImageIO;
import javax.mail.MessagingException;
import javax.mail.internet.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.ConnectException;
import java.net.NoRouteToHostException;
import java.net.UnknownHostException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Exchange session through Outlook Web Access (DAV)
 */
public abstract class ExchangeSession {

    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    /**
     * Reference GMT timezone to format dates
     */
    public static final SimpleTimeZone GMT_TIMEZONE = new SimpleTimeZone(0, "GMT");

    protected static final Set<String> USER_NAME_FIELDS = new HashSet<String>();

    static {
        USER_NAME_FIELDS.add("username");
        USER_NAME_FIELDS.add("txtUserName");
        USER_NAME_FIELDS.add("userid");
        USER_NAME_FIELDS.add("SafeWordUser");
        USER_NAME_FIELDS.add("user_name");
    }

    protected static final Set<String> PASSWORD_FIELDS = new HashSet<String>();

    static {
        PASSWORD_FIELDS.add("password");
        PASSWORD_FIELDS.add("txtUserPass");
        PASSWORD_FIELDS.add("pw");
        PASSWORD_FIELDS.add("basicPassword");
    }

    protected static final Set<String> TOKEN_FIELDS = new HashSet<String>();

    static {
        TOKEN_FIELDS.add("SafeWordPassword");
        TOKEN_FIELDS.add("passcode");
    }

    public static final String PUBLIC_ROOT = "/public/";
    public static final String CALENDAR = "calendar";
    public static final String TASKS = "tasks";
    /**
     * Contacts folder logical name
     */
    public static final String CONTACTS = "contacts";
    public static final String ADDRESSBOOK = "addressbook";
    public static final String INBOX = "INBOX";
    public static final String LOWER_CASE_INBOX = "inbox";
    public static final String SENT = "Sent";
    public static final String SENDMSG = "##DavMailSubmissionURI##";
    public static final String DRAFTS = "Drafts";
    public static final String TRASH = "Trash";
    public static final String JUNK = "Junk";
    public static final String UNSENT = "Unsent Messages";

    static {
        // Adjust Mime decoder settings
        System.setProperty("mail.mime.ignoreunknownencoding", "true");
        System.setProperty("mail.mime.decodetext.strict", "false");
    }

    protected String publicFolderUrl;

    /**
     * Base user mailboxes path (used to select folder)
     */
    public String mailPath;
    public String rootPath;
    public String email;
    public String alias;
    /**
     * Lower case Caldav path to current user mailbox.
     * /users/<i>email</i>
     */
    protected String currentMailboxPath;
    protected final HttpClient httpClient;

    protected String userName;
    /**
     * A OTP pre-auth page may require a different username.
     */
    private String preAuthUsername;

    protected ExchangeVersion serverVersion;

    protected static final String YYYY_MM_DD_HH_MM_SS = "yyyy/MM/dd HH:mm:ss";
    private static final String YYYYMMDD_T_HHMMSS_Z = "yyyyMMdd'T'HHmmss'Z'";
    protected static final String YYYY_MM_DD_T_HHMMSS_Z = "yyyy-MM-dd'T'HH:mm:ss'Z'";
    private static final String YYYY_MM_DD = "yyyy-MM-dd";
    private static final String YYYY_MM_DD_T_HHMMSS_SSS_Z = "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'";

    /**
     * Logon form user name fields.
     */
    private final List<String> userNameInputs = new ArrayList<String>();
    /**
     * Logon form password field, default is password.
     */
    private String passwordInput = null;
    /**
     * Tells if, during the login navigation, an OTP pre-auth page has been found.
     */
    private boolean otpPreAuthFound = false;
    /**
     * Lets the user try again a couple of times to enter the OTP pre-auth key before giving up.
     */
    private int otpPreAuthRetries = 0;
    /**
     * Maximum number of times the user can try to input again the OTP pre-auth key before giving up.
     */
    private static final int MAX_OTP_RETRIES = 3;

    /**
     * Create an exchange session for the given URL.
     * The session is established for given userName and password
     *
     * @param url      Exchange url
     * @param userName user login name
     * @param password user password
     * @throws IOException on error
     */
    public ExchangeSession(String url, String userName, String password) throws IOException {
        this.userName = userName;
        try {
            httpClient = DavGatewayHttpClientFacade.getInstance(url);
            // set private connection pool
            DavGatewayHttpClientFacade.createMultiThreadedHttpConnectionManager(httpClient);
            boolean isBasicAuthentication = isBasicAuthentication(httpClient, url);
            // clear cookies created by authentication test
            httpClient.getState().clearCookies();

            // The user may have configured an OTP pre-auth username. It is processed
            // so early because OTP pre-auth may disappear in the Exchange LAN and this
            // helps the user to not change is account settings in mail client at each network change.
            if (preAuthUsername == null) {
                // Searches for the delimiter in configured username for the pre-auth user. 
                // The double-quote is not allowed inside email addresses anyway.
                int doubleQuoteIndex = this.userName.indexOf('"');
                if (doubleQuoteIndex > 0) {
                    preAuthUsername = this.userName.substring(0, doubleQuoteIndex);
                    this.userName = this.userName.substring(doubleQuoteIndex + 1);
                } else {
                    // No doublequote: the pre-auth user is the full username, or it is not used at all.
                    preAuthUsername = this.userName;
                }
            }

            DavGatewayHttpClientFacade.setCredentials(httpClient, userName, password);

            // get webmail root url
            // providing credentials
            // manually follow redirect
            HttpMethod method = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, url);

            if (!this.isAuthenticated()) {
                if (isBasicAuthentication) {
                    int status = method.getStatusCode();

                    if (status == HttpStatus.SC_UNAUTHORIZED) {
                        method.releaseConnection();
                        throw new DavMailAuthenticationException("EXCEPTION_AUTHENTICATION_FAILED");
                    } else if (status != HttpStatus.SC_OK) {
                        method.releaseConnection();
                        throw DavGatewayHttpClientFacade.buildHttpException(method);
                    }
                    // workaround for basic authentication on /exchange and form based authentication at /owa
                    if ("/owa/auth/logon.aspx".equals(method.getPath())) {
                        method = formLogin(httpClient, method, userName, password);
                    }
                } else {
                    method = formLogin(httpClient, method, userName, password);
                }
            }

            // avoid 401 roundtrips, only if NTLM is disabled and basic authentication enabled
            if (isBasicAuthentication && !DavGatewayHttpClientFacade.hasNTLMorNegotiate(httpClient)) {
                httpClient.getParams().setParameter(HttpClientParams.PREEMPTIVE_AUTHENTICATION, true);
            }

            buildSessionInfo(method);

        } catch (DavMailAuthenticationException exc) {
            LOGGER.error(exc.getMessage());
            throw exc;
        } catch (ConnectException exc) {
            BundleMessage message = new BundleMessage("EXCEPTION_CONNECT", exc.getClass().getName(), exc.getMessage());
            ExchangeSession.LOGGER.error(message);
            throw new DavMailException("EXCEPTION_DAVMAIL_CONFIGURATION", message);
        } catch (UnknownHostException exc) {
            BundleMessage message = new BundleMessage("EXCEPTION_CONNECT", exc.getClass().getName(), exc.getMessage());
            ExchangeSession.LOGGER.error(message);
            throw new DavMailException("EXCEPTION_DAVMAIL_CONFIGURATION", message);
        } catch (WebdavNotAvailableException exc) {
            throw exc;
        } catch (IOException exc) {
            LOGGER.error(BundleMessage.formatLog("EXCEPTION_EXCHANGE_LOGIN_FAILED", exc));
            throw new DavMailException("EXCEPTION_EXCHANGE_LOGIN_FAILED", exc);
        }
        LOGGER.debug("Session " + this + " created");
    }

    /**
     * Format date to exchange search format.
     *
     * @param date date object
     * @return formatted search date
     */
    public abstract String formatSearchDate(Date date);

    /**
     * Return standard zulu date formatter.
     *
     * @return zulu date formatter
     */
    public static SimpleDateFormat getZuluDateFormat() {
        SimpleDateFormat dateFormat = new SimpleDateFormat(YYYYMMDD_T_HHMMSS_Z, Locale.ENGLISH);
        dateFormat.setTimeZone(GMT_TIMEZONE);
        return dateFormat;
    }

    protected static SimpleDateFormat getVcardBdayFormat() {
        SimpleDateFormat dateFormat = new SimpleDateFormat(YYYY_MM_DD, Locale.ENGLISH);
        dateFormat.setTimeZone(GMT_TIMEZONE);
        return dateFormat;
    }

    protected static SimpleDateFormat getExchangeZuluDateFormat() {
        SimpleDateFormat dateFormat = new SimpleDateFormat(YYYY_MM_DD_T_HHMMSS_Z, Locale.ENGLISH);
        dateFormat.setTimeZone(GMT_TIMEZONE);
        return dateFormat;
    }

    protected static SimpleDateFormat getExchangeZuluDateFormatMillisecond() {
        SimpleDateFormat dateFormat = new SimpleDateFormat(YYYY_MM_DD_T_HHMMSS_SSS_Z, Locale.ENGLISH);
        dateFormat.setTimeZone(GMT_TIMEZONE);
        return dateFormat;
    }

    protected static Date parseDate(String dateString) throws ParseException {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
        dateFormat.setTimeZone(GMT_TIMEZONE);
        return dateFormat.parse(dateString);
    }


    /**
     * Test if the session expired.
     *
     * @return true this session expired
     * @throws NoRouteToHostException on error
     * @throws UnknownHostException   on error
     */
    public boolean isExpired() throws NoRouteToHostException, UnknownHostException {
        boolean isExpired = false;
        try {
            getFolder("");
        } catch (UnknownHostException exc) {
            throw exc;
        } catch (NoRouteToHostException exc) {
            throw exc;
        } catch (IOException e) {
            isExpired = true;
        }

        return isExpired;
    }

    /**
     * Test authentication mode : form based or basic.
     *
     * @param url        exchange base URL
     * @param httpClient httpClient instance
     * @return true if basic authentication detected
     */
    protected boolean isBasicAuthentication(HttpClient httpClient, String url) {
        return DavGatewayHttpClientFacade.getHttpStatus(httpClient, url) == HttpStatus.SC_UNAUTHORIZED;
    }

    protected String getAbsoluteUri(HttpMethod method, String path) throws URIException {
        URI uri = method.getURI();
        if (path != null) {
            // reset query string
            uri.setQuery(null);
            if (path.startsWith("/")) {
                // path is absolute, replace method path
                uri.setPath(path);
            } else if (path.startsWith("http://") || path.startsWith("https://")) {
                return path;
            } else {
                // relative path, build new path
                String currentPath = method.getPath();
                int end = currentPath.lastIndexOf('/');
                if (end >= 0) {
                    uri.setPath(currentPath.substring(0, end + 1) + path);
                } else {
                    throw new URIException(uri.getURI());
                }
            }
        }
        return uri.getURI();
    }

    protected String getScriptBasedFormURL(HttpMethod initmethod, String pathQuery) throws URIException {
        URI initmethodURI = initmethod.getURI();
        int queryIndex = pathQuery.indexOf('?');
        if (queryIndex >= 0) {
            if (queryIndex > 0) {
                // update path
                String newPath = pathQuery.substring(0, queryIndex);
                if (newPath.startsWith("/")) {
                    // absolute path
                    initmethodURI.setPath(newPath);
                } else {
                    String currentPath = initmethodURI.getPath();
                    int folderIndex = currentPath.lastIndexOf('/');
                    if (folderIndex >= 0) {
                        // replace relative path
                        initmethodURI.setPath(currentPath.substring(0, folderIndex + 1) + newPath);
                    } else {
                        // should not happen
                        initmethodURI.setPath('/' + newPath);
                    }
                }
            }
            initmethodURI.setQuery(pathQuery.substring(queryIndex + 1));
        }
        return initmethodURI.getURI();
    }

    /**
     * Try to find logon method path from logon form body.
     *
     * @param httpClient httpClient instance
     * @param initmethod form body http method
     * @return logon method
     * @throws IOException on error
     */
    protected HttpMethod buildLogonMethod(HttpClient httpClient, HttpMethod initmethod) throws IOException {

        HttpMethod logonMethod = null;

        // create an instance of HtmlCleaner
        HtmlCleaner cleaner = new HtmlCleaner();

        // A OTP token authentication form in a previous page could have username fields with different names
        userNameInputs.clear();

        try {
            TagNode node = cleaner.clean(initmethod.getResponseBodyAsStream());
            List forms = node.getElementListByName("form", true);
            TagNode logonForm = null;
            // select form
            if (forms.size() == 1) {
                logonForm = (TagNode) forms.get(0);
            } else if (forms.size() > 1) {
                for (Object form : forms) {
                    if ("logonForm".equals(((TagNode) form).getAttributeByName("name"))) {
                        logonForm = ((TagNode) form);
                    }
                }
            }
            if (logonForm != null) {
                String logonMethodPath = logonForm.getAttributeByName("action");

                // workaround for broken form with empty action
                if (logonMethodPath != null && logonMethodPath.length() == 0) {
                    logonMethodPath = "/owa/auth.owa";
                }

                logonMethod = new PostMethod(getAbsoluteUri(initmethod, logonMethodPath));

                // retrieve lost inputs attached to body
                List inputList = node.getElementListByName("input", true);

                for (Object input : inputList) {
                    String type = ((TagNode) input).getAttributeByName("type");
                    String name = ((TagNode) input).getAttributeByName("name");
                    String value = ((TagNode) input).getAttributeByName("value");
                    if ("hidden".equalsIgnoreCase(type) && name != null && value != null) {
                        ((PostMethod) logonMethod).addParameter(name, value);
                    }
                    // custom login form
                    if (USER_NAME_FIELDS.contains(name)) {
                        userNameInputs.add(name);
                    } else if (PASSWORD_FIELDS.contains(name)) {
                        passwordInput = name;
                    } else if ("addr".equals(name)) {
                        // this is not a logon form but a redirect form
                        HttpMethod newInitMethod = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, logonMethod);
                        logonMethod = buildLogonMethod(httpClient, newInitMethod);
                    } else if (TOKEN_FIELDS.contains(name)) {
                        // one time password, ask it to the user
                        ((PostMethod) logonMethod).addParameter(name, DavGatewayOTPPrompt.getOneTimePassword());
                    } else if ("otc".equals(name)) {
                        // captcha image, get image and ask user
                        String pinsafeUser = getAliasFromLogin();
                        if (pinsafeUser == null) {
                            pinsafeUser = userName;
                        }
                        GetMethod getMethod = new GetMethod("/PINsafeISAFilter.dll?username=" + pinsafeUser);
                        try {
                            int status = httpClient.executeMethod(getMethod);
                            if (status != HttpStatus.SC_OK) {
                                throw DavGatewayHttpClientFacade.buildHttpException(getMethod);
                            }
                            BufferedImage captchaImage = ImageIO.read(getMethod.getResponseBodyAsStream());
                            ((PostMethod) logonMethod).addParameter(name, DavGatewayOTPPrompt.getCaptchaValue(captchaImage));

                        } finally {
                            getMethod.releaseConnection();
                        }
                    }
                }
            } else {
                List frameList = node.getElementListByName("frame", true);
                if (frameList.size() == 1) {
                    String src = ((TagNode) frameList.get(0)).getAttributeByName("src");
                    if (src != null) {
                        LOGGER.debug("Frames detected in form page, try frame content");
                        initmethod.releaseConnection();
                        HttpMethod newInitMethod = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, src);
                        logonMethod = buildLogonMethod(httpClient, newInitMethod);
                    }
                } else {
                    // another failover for script based logon forms (Exchange 2007)
                    List scriptList = node.getElementListByName("script", true);
                    for (Object script : scriptList) {
                        List contents = ((TagNode) script).getChildren();
                        for (Object content : contents) {
                            if (content instanceof CommentNode) {
                                String scriptValue = ((CommentNode) content).getCommentedContent();
                                String sUrl = StringUtil.getToken(scriptValue, "var a_sUrl = \"", "\"");
                                String sLgn = StringUtil.getToken(scriptValue, "var a_sLgnQS = \"", "\"");
                                if (sLgn == null) {
                                    sLgn = StringUtil.getToken(scriptValue, "var a_sLgn = \"", "\"");
                                }
                                if (sUrl != null && sLgn != null) {
                                    String src = getScriptBasedFormURL(initmethod, sLgn + sUrl);
                                    LOGGER.debug("Detected script based logon, redirect to form at " + src);
                                    HttpMethod newInitMethod = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, src);
                                    logonMethod = buildLogonMethod(httpClient, newInitMethod);
                                }

                            } else if (content instanceof ContentNode) {
                                // Microsoft Forefront Unified Access Gateway redirect
                                String scriptValue = ((ContentNode) content).getContent().toString();
                                String location = StringUtil.getToken(scriptValue, "window.location.replace(\"", "\"");
                                if (location != null) {
                                    LOGGER.debug("Post logon redirect to: " + location);
                                    logonMethod = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, location);
                                }
                            }
                        }
                    }
                }
            }
        } catch (IOException e) {
            LOGGER.error("Error parsing login form at " + initmethod.getURI());
        } finally {
            initmethod.releaseConnection();
        }

        return logonMethod;
    }

    protected HttpMethod postLogonMethod(HttpClient httpClient, HttpMethod logonMethod, String userName, String password) throws IOException {

        setAuthFormFields((PostMethod)logonMethod, httpClient, password);

        // add exchange 2010 PBack cookie in compatibility mode
        httpClient.getState().addCookie(new Cookie(httpClient.getHostConfiguration().getHost(), "PBack", "0", "/", null, false));

        logonMethod = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, logonMethod);

        // test form based authentication
        checkFormLoginQueryString(logonMethod);

        // workaround for post logon script redirect
        if (!isAuthenticated()) {
            // try to get new method from script based redirection
            logonMethod = buildLogonMethod(httpClient, logonMethod);

            if (logonMethod != null) {
                if (otpPreAuthFound && otpPreAuthRetries < MAX_OTP_RETRIES) {
                    // A OTP pre-auth page has been found, it is needed to restart the login process.
                    // This applies to both the case the user entered a good OTP code (the usual login process
                    // takes place) and the case the user entered a wrong OTP code (another code will be asked to him).
                    // The user has up to MAX_OTP_RETRIES chances to input a valid OTP key.
                    return postLogonMethod(httpClient, logonMethod, userName, password);
                }

                // if logonMethod is not null, try to follow redirection
                logonMethod = DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, logonMethod);
                checkFormLoginQueryString(logonMethod);
                // also check cookies
                if (!isAuthenticated()) {
                    throwAuthenticationFailed();
                }
            } else {
                // authentication failed
                throwAuthenticationFailed();
            }
        }

        // check for language selection form
        if (logonMethod != null && "/owa/languageselection.aspx".equals(logonMethod.getPath())) {
            // need to submit form
            logonMethod = submitLanguageSelectionForm(logonMethod);
        }
        return logonMethod;
    }

    protected void setAuthFormFields(PostMethod logonMethod, HttpClient httpClient, String password) throws IllegalArgumentException {
        String userNameInput;
        if (userNameInputs.size() == 2) {
            String userid;
            // multiple username fields, split userid|username on |
            int pipeIndex = userName.indexOf('|');
            if (pipeIndex < 0) {
                LOGGER.debug("Multiple user fields detected, please use userid|username as user name in client, except when userid is username");
                userid = userName;
            } else {
                userid = userName.substring(0, pipeIndex);
                userName = userName.substring(pipeIndex + 1);
                // adjust credentials
                DavGatewayHttpClientFacade.setCredentials(httpClient, userName, password);
            }
            logonMethod.removeParameter("userid");
            logonMethod.addParameter("userid", userid);

            userNameInput = "username";
        } else if (userNameInputs.size() == 1) {
            // simple username field
            userNameInput = userNameInputs.get(0);
        } else {
            // should not happen
            userNameInput = "username";
        }
        // make sure username and password fields are empty
        logonMethod.removeParameter(userNameInput);
        if (passwordInput != null) {
            logonMethod.removeParameter(passwordInput);
        }
        logonMethod.removeParameter("trusted");
        logonMethod.removeParameter("flags");

        if (passwordInput == null) {
            // This is a OTP pre-auth page. A different username may be required.
            otpPreAuthFound = true;
            otpPreAuthRetries++;
            logonMethod.addParameter(userNameInput, preAuthUsername);
        } else {
            otpPreAuthFound = false;
            otpPreAuthRetries = 0;
            // This is a regular Exchange login page
            logonMethod.addParameter(userNameInput, userName);
            logonMethod.addParameter(passwordInput, password);
            logonMethod.addParameter("trusted", "4");
            logonMethod.addParameter("flags", "4");
        }
    }

    protected HttpMethod formLogin(HttpClient httpClient, HttpMethod initmethod, String userName, String password) throws IOException {
        LOGGER.debug("Form based authentication detected");

        HttpMethod logonMethod = buildLogonMethod(httpClient, initmethod);
        if (logonMethod == null) {
            LOGGER.debug("Authentication form not found at " + initmethod.getURI() + ", trying default url");
            logonMethod = new PostMethod("/owa/auth/owaauth.dll");
        }
        return postLogonMethod(httpClient, logonMethod, userName, password);
    }

    protected HttpMethod submitLanguageSelectionForm(HttpMethod logonMethod) throws IOException {
        PostMethod postLanguageFormMethod;
        // create an instance of HtmlCleaner
        HtmlCleaner cleaner = new HtmlCleaner();

        try {
            TagNode node = cleaner.clean(logonMethod.getResponseBodyAsStream());
            List forms = node.getElementListByName("form", true);
            TagNode languageForm;
            // select form
            if (forms.size() == 1) {
                languageForm = (TagNode) forms.get(0);
            } else {
                throw new IOException("Form not found");
            }
            String languageMethodPath = languageForm.getAttributeByName("action");

            postLanguageFormMethod = new PostMethod(getAbsoluteUri(logonMethod, languageMethodPath));

            List inputList = languageForm.getElementListByName("input", true);
            for (Object input : inputList) {
                String name = ((TagNode) input).getAttributeByName("name");
                String value = ((TagNode) input).getAttributeByName("value");
                if (name != null && value != null) {
                    postLanguageFormMethod.addParameter(name, value);
                }
            }
            List selectList = languageForm.getElementListByName("select", true);
            for (Object select : selectList) {
                String name = ((TagNode) select).getAttributeByName("name");
                List optionList = ((TagNode) select).getElementListByName("option", true);
                String value = null;
                for (Object option : optionList) {
                    if (((TagNode) option).getAttributeByName("selected") != null) {
                        value = ((TagNode) option).getAttributeByName("value");
                        break;
                    }
                }
                if (name != null && value != null) {
                    postLanguageFormMethod.addParameter(name, value);
                }
            }
        } catch (IOException e) {
            String errorMessage = "Error parsing language selection form at " + logonMethod.getURI();
            LOGGER.error(errorMessage);
            throw new IOException(errorMessage);
        } finally {
            logonMethod.releaseConnection();
        }

        return DavGatewayHttpClientFacade.executeFollowRedirects(httpClient, postLanguageFormMethod);
    }

    /**
     * Look for session cookies.
     *
     * @return true if session cookies are available
     */
    protected boolean isAuthenticated() {
        boolean authenticated = false;
        for (Cookie cookie : httpClient.getState().getCookies()) {
            // Exchange 2003 cookies
            if (cookie.getName().startsWith("cadata") || "sessionid".equals(cookie.getName())
                    // Exchange 2007 cookie
                    || "UserContext".equals(cookie.getName())
                    // Direct EWS access
                    || "exchangecookie".equals(cookie.getName())
                    ) {
                authenticated = true;
                break;
            }
        }
        return authenticated;
    }

    protected void checkFormLoginQueryString(HttpMethod logonMethod) throws DavMailAuthenticationException {
        String queryString = logonMethod.getQueryString();
        if (queryString != null && (queryString.contains("reason=2") || queryString.contains("reason=4"))) {
            logonMethod.releaseConnection();
            throwAuthenticationFailed();
        }
    }

    protected void throwAuthenticationFailed() throws DavMailAuthenticationException {
        if (this.userName != null && this.userName.contains("\\")) {
            throw new DavMailAuthenticationException("EXCEPTION_AUTHENTICATION_FAILED");
        } else {
            throw new DavMailAuthenticationException("EXCEPTION_AUTHENTICATION_FAILED_RETRY");
        }
    }

    protected abstract void buildSessionInfo(HttpMethod method) throws DavMailException;

    /**
     * Create message in specified folder.
     * Will overwrite an existing message with same subject in the same folder
     *
     * @param folderPath  Exchange folder path
     * @param messageName message name
     * @param properties  message properties (flags)
     * @param mimeMessage MIME message
     * @throws IOException when unable to create message
     */
    public abstract void createMessage(String folderPath, String messageName, HashMap<String, String> properties, MimeMessage mimeMessage) throws IOException;

    /**
     * Update given properties on message.
     *
     * @param message    Exchange message
     * @param properties Webdav properties map
     * @throws IOException on error
     */
    public abstract void updateMessage(Message message, Map<String, String> properties) throws IOException;


    /**
     * Delete Exchange message.
     *
     * @param message Exchange message
     * @throws IOException on error
     */
    public abstract void deleteMessage(Message message) throws IOException;

    /**
     * Get raw MIME message content
     *
     * @param message Exchange message
     * @return message body
     * @throws IOException on error
     */
    public abstract byte[] getContent(Message message) throws IOException;

    protected static final Set<String> POP_MESSAGE_ATTRIBUTES = new HashSet<String>();

    static {
        POP_MESSAGE_ATTRIBUTES.add("uid");
        POP_MESSAGE_ATTRIBUTES.add("imapUid");
        POP_MESSAGE_ATTRIBUTES.add("messageSize");
    }

    /**
     * Return folder message list with id and size only (for POP3 listener).
     *
     * @param folderName Exchange folder name
     * @return folder message list
     * @throws IOException on error
     */
    public MessageList getAllMessageUidAndSize(String folderName) throws IOException {
        return searchMessages(folderName, POP_MESSAGE_ATTRIBUTES, null);
    }

    protected static final Set<String> IMAP_MESSAGE_ATTRIBUTES = new HashSet<String>();

    static {
        IMAP_MESSAGE_ATTRIBUTES.add("permanenturl");
        IMAP_MESSAGE_ATTRIBUTES.add("urlcompname");
        IMAP_MESSAGE_ATTRIBUTES.add("uid");
        IMAP_MESSAGE_ATTRIBUTES.add("messageSize");
        IMAP_MESSAGE_ATTRIBUTES.add("imapUid");
        IMAP_MESSAGE_ATTRIBUTES.add("junk");
        IMAP_MESSAGE_ATTRIBUTES.add("flagStatus");
        IMAP_MESSAGE_ATTRIBUTES.add("messageFlags");
        IMAP_MESSAGE_ATTRIBUTES.add("lastVerbExecuted");
        IMAP_MESSAGE_ATTRIBUTES.add("read");
        IMAP_MESSAGE_ATTRIBUTES.add("deleted");
        IMAP_MESSAGE_ATTRIBUTES.add("date");
        IMAP_MESSAGE_ATTRIBUTES.add("lastmodified");
        // OSX IMAP requests content-class
        IMAP_MESSAGE_ATTRIBUTES.add("contentclass");
        IMAP_MESSAGE_ATTRIBUTES.add("keywords");
    }

    protected static final Set<String> UID_MESSAGE_ATTRIBUTES = new HashSet<String>();

    static {
        UID_MESSAGE_ATTRIBUTES.add("uid");
    }

    /**
     * Get all folder messages.
     *
     * @param folderPath Exchange folder name
     * @return message list
     * @throws IOException on error
     */
    public MessageList searchMessages(String folderPath) throws IOException {
        return searchMessages(folderPath, IMAP_MESSAGE_ATTRIBUTES, null);
    }

    /**
     * Search folder for messages matching conditions, with attributes needed by IMAP listener.
     *
     * @param folderName Exchange folder name
     * @param condition  search filter
     * @return message list
     * @throws IOException on error
     */
    public MessageList searchMessages(String folderName, Condition condition) throws IOException {
        return searchMessages(folderName, IMAP_MESSAGE_ATTRIBUTES, condition);
    }

    /**
     * Search folder for messages matching conditions, with given attributes.
     *
     * @param folderName Exchange folder name
     * @param attributes requested Webdav attributes
     * @param condition  search filter
     * @return message list
     * @throws IOException on error
     */
    public abstract MessageList searchMessages(String folderName, Set<String> attributes, Condition condition) throws IOException;

    /**
     * Get server version (Exchange2003, Exchange2007 or Exchange2010)
     *
     * @return server version
     */
    public ExchangeVersion getServerVersion() {
        return serverVersion;
    }

    @SuppressWarnings({"JavaDoc"})
    public enum Operator {
        Or, And, Not, IsEqualTo,
        IsGreaterThan, IsGreaterThanOrEqualTo,
        IsLessThan, IsLessThanOrEqualTo,
        IsNull, IsTrue, IsFalse,
        Like, StartsWith, Contains
    }

    /**
     * And search filter.
     *
     * @param condition search conditions
     * @return condition
     */
    public abstract MultiCondition and(Condition... condition);

    /**
     * Or search filter.
     *
     * @param condition search conditions
     * @return condition
     */
    public abstract MultiCondition or(Condition... condition);

    /**
     * Not search filter.
     *
     * @param condition search condition
     * @return condition
     */
    public abstract Condition not(Condition condition);

    /**
     * Equals condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition isEqualTo(String attributeName, String value);

    /**
     * Equals condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition isEqualTo(String attributeName, int value);

    /**
     * MIME header equals condition.
     *
     * @param headerName MIME header name
     * @param value      attribute value
     * @return condition
     */
    public abstract Condition headerIsEqualTo(String headerName, String value);

    /**
     * Greater than or equals condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition gte(String attributeName, String value);

    /**
     * Greater than condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition gt(String attributeName, String value);

    /**
     * Lower than condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition lt(String attributeName, String value);

    /**
     * Lower than or equals condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    @SuppressWarnings({"UnusedDeclaration"})
    public abstract Condition lte(String attributeName, String value);

    /**
     * Contains condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition contains(String attributeName, String value);

    /**
     * Starts with condition.
     *
     * @param attributeName logical Exchange attribute name
     * @param value         attribute value
     * @return condition
     */
    public abstract Condition startsWith(String attributeName, String value);

    /**
     * Is null condition.
     *
     * @param attributeName logical Exchange attribute name
     * @return condition
     */
    public abstract Condition isNull(String attributeName);

    /**
     * Is true condition.
     *
     * @param attributeName logical Exchange attribute name
     * @return condition
     */
    public abstract Condition isTrue(String attributeName);

    /**
     * Is false condition.
     *
     * @param attributeName logical Exchange attribute name
     * @return condition
     */
    public abstract Condition isFalse(String attributeName);

    /**
     * Search mail and generic folders under given folder.
     * Exclude calendar and contacts folders
     *
     * @param folderName Exchange folder name
     * @param recursive  deep search if true
     * @return list of folders
     * @throws IOException on error
     */
    public List<Folder> getSubFolders(String folderName, boolean recursive) throws IOException {
        Condition folderCondition = null;
        if (!Settings.getBooleanProperty("davmail.imapIncludeSpecialFolders", false)) {
            folderCondition = or(isEqualTo("folderclass", "IPF.Note"), isNull("folderclass"));
        }
        List<Folder> results = getSubFolders(folderName, folderCondition,
                recursive);
        // need to include base folder in recursive search, except on root
        if (recursive && folderName.length() > 0) {
            results.add(getFolder(folderName));
        }

        return results;
    }

    /**
     * Search calendar folders under given folder.
     *
     * @param folderName Exchange folder name
     * @param recursive  deep search if true
     * @return list of folders
     * @throws IOException on error
     */
    public List<Folder> getSubCalendarFolders(String folderName, boolean recursive) throws IOException {
        return getSubFolders(folderName, isEqualTo("folderclass", "IPF.Appointment"), recursive);
    }

    /**
     * Search folders under given folder matching filter.
     *
     * @param folderName Exchange folder name
     * @param condition  search filter
     * @param recursive  deep search if true
     * @return list of folders
     * @throws IOException on error
     */
    public abstract List<Folder> getSubFolders(String folderName, Condition condition, boolean recursive) throws IOException;

    /**
     * Delete oldest messages in trash.
     * keepDelay is the number of days to keep messages in trash before delete
     *
     * @throws IOException when unable to purge messages
     */
    public void purgeOldestTrashAndSentMessages() throws IOException {
        int keepDelay = Settings.getIntProperty("davmail.keepDelay");
        if (keepDelay != 0) {
            purgeOldestFolderMessages(TRASH, keepDelay);
        }
        // this is a new feature, default is : do nothing
        int sentKeepDelay = Settings.getIntProperty("davmail.sentKeepDelay");
        if (sentKeepDelay != 0) {
            purgeOldestFolderMessages(SENT, sentKeepDelay);
        }
    }

    protected void purgeOldestFolderMessages(String folderPath, int keepDelay) throws IOException {
        Calendar cal = Calendar.getInstance();
        cal.add(Calendar.DAY_OF_MONTH, -keepDelay);
        LOGGER.debug("Delete messages in " + folderPath + " not modified since " + cal.getTime());

        MessageList messages = searchMessages(folderPath, UID_MESSAGE_ATTRIBUTES,
                lt("lastmodified", formatSearchDate(cal.getTime())));

        for (Message message : messages) {
            message.delete();
        }
    }

    protected void convertResentHeader(MimeMessage mimeMessage, String headerName) throws MessagingException {
        String[] resentHeader = mimeMessage.getHeader("Resent-" + headerName);
        if (resentHeader != null) {
            mimeMessage.removeHeader("Resent-" + headerName);
            mimeMessage.removeHeader(headerName);
            for (String value : resentHeader) {
                mimeMessage.addHeader(headerName, value);
            }
        }
    }

    protected String lastSentMessageId;

    /**
     * Send message in reader to recipients.
     * Detect visible recipients in message body to determine bcc recipients
     *
     * @param rcptToRecipients recipients list
     * @param mimeMessage      mime message
     * @throws IOException        on error
     * @throws MessagingException on error
     */
    public void sendMessage(List<String> rcptToRecipients, MimeMessage mimeMessage) throws IOException, MessagingException {
        // detect duplicate send command
        String messageId = mimeMessage.getMessageID();
        if (lastSentMessageId != null && lastSentMessageId.equals(messageId)) {
            LOGGER.debug("Dropping message id " + messageId + ": already sent");
            return;
        }
        lastSentMessageId = messageId;

        convertResentHeader(mimeMessage, "From");
        convertResentHeader(mimeMessage, "To");
        convertResentHeader(mimeMessage, "Cc");
        convertResentHeader(mimeMessage, "Bcc");
        convertResentHeader(mimeMessage, "Message-Id");

        // do not allow send as another user on Exchange 2003
        if (serverVersion.isExchange2003() || Settings.getBooleanProperty("davmail.smtpStripFrom", false)) {
            mimeMessage.removeHeader("From");
        }

        // remove visible recipients from list
        Set<String> visibleRecipients = new HashSet<String>();
        List<InternetAddress> recipients = getAllRecipients(mimeMessage);
        for (InternetAddress address : recipients) {
            visibleRecipients.add((address.getAddress().toLowerCase()));
        }
        for (String recipient : rcptToRecipients) {
            if (!visibleRecipients.contains(recipient.toLowerCase())) {
                mimeMessage.addRecipient(javax.mail.Message.RecipientType.BCC, new InternetAddress(recipient));
            }
        }
        sendMessage(mimeMessage);

    }

    static final String[] RECIPIENT_HEADERS = {"to", "cc", "bcc"};

    protected List<InternetAddress> getAllRecipients(MimeMessage mimeMessage) throws MessagingException {
        List<InternetAddress> recipientList = new ArrayList<InternetAddress>();
        for (String recipientHeader : RECIPIENT_HEADERS) {
            final String recipientHeaderValue = mimeMessage.getHeader(recipientHeader, ",");
            if (recipientHeaderValue != null) {
                // parse headers in non strict mode
                recipientList.addAll(Arrays.asList(InternetAddress.parseHeader(recipientHeaderValue, false)));
            }

        }
        return recipientList;
    }

    /**
     * Send Mime message.
     *
     * @param mimeMessage MIME message
     * @throws IOException        on error
     * @throws MessagingException on error
     */
    public abstract void sendMessage(MimeMessage mimeMessage) throws IOException, MessagingException;

    /**
     * Get folder object.
     * Folder name can be logical names INBOX, Drafts, Trash or calendar,
     * or a path relative to user base folder or absolute path.
     *
     * @param folderPath folder path
     * @return Folder object
     * @throws IOException on error
     */
    public Folder getFolder(String folderPath) throws IOException {
        Folder folder = internalGetFolder(folderPath);
        if (isMainCalendar(folderPath)) {
            Folder taskFolder = internalGetFolder(TASKS);
            folder.ctag += taskFolder.ctag;
        }
        return folder;
    }

    protected abstract Folder internalGetFolder(String folderName) throws IOException;

    /**
     * Check folder ctag and reload messages as needed.
     *
     * @param currentFolder current folder
     * @return true if folder changed
     * @throws IOException on error
     */
    public boolean refreshFolder(Folder currentFolder) throws IOException {
        Folder newFolder = getFolder(currentFolder.folderPath);
        if (currentFolder.ctag == null || !currentFolder.ctag.equals(newFolder.ctag)) {
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug("Contenttag changed on " + currentFolder.folderPath + ' '
                        + currentFolder.ctag + " => " + newFolder.ctag + ", reloading messages");
            }
            currentFolder.hasChildren = newFolder.hasChildren;
            currentFolder.noInferiors = newFolder.noInferiors;
            currentFolder.unreadCount = newFolder.unreadCount;
            currentFolder.ctag = newFolder.ctag;
            currentFolder.etag = newFolder.etag;
            if (newFolder.uidNext > currentFolder.uidNext) {
                currentFolder.uidNext = newFolder.uidNext;
            }
            currentFolder.loadMessages();
            return true;
        } else {
            return false;
        }
    }

    /**
     * Create Exchange message folder.
     *
     * @param folderName logical folder name
     * @return status
     * @throws IOException on error
     */
    public int createMessageFolder(String folderName) throws IOException {
        return createFolder(folderName, "IPF.Note", null);
    }

    /**
     * Create Exchange calendar folder.
     *
     * @param folderName logical folder name
     * @param properties folder properties
     * @return status
     * @throws IOException on error
     */
    public int createCalendarFolder(String folderName, Map<String, String> properties) throws IOException {
        return createFolder(folderName, "IPF.Appointment", properties);
    }

    /**
     * Create Exchange contact folder.
     *
     * @param folderName logical folder name
     * @param properties folder properties
     * @return status
     * @throws IOException on error
     */
    public int createContactFolder(String folderName, Map<String, String> properties) throws IOException {
        return createFolder(folderName, "IPF.Contact", properties);
    }

    /**
     * Create Exchange folder with given folder class.
     *
     * @param folderName  logical folder name
     * @param folderClass folder class
     * @param properties  folder properties
     * @return status
     * @throws IOException on error
     */
    public abstract int createFolder(String folderName, String folderClass, Map<String, String> properties) throws IOException;

    /**
     * Update Exchange folder properties.
     *
     * @param folderName logical folder name
     * @param properties folder properties
     * @return status
     * @throws IOException on error
     */
    public abstract int updateFolder(String folderName, Map<String, String> properties) throws IOException;

    /**
     * Delete Exchange folder.
     *
     * @param folderName logical folder name
     * @throws IOException on error
     */
    public abstract void deleteFolder(String folderName) throws IOException;

    /**
     * Copy message to target folder
     *
     * @param message      Exchange message
     * @param targetFolder target folder
     * @throws IOException on error
     */
    public abstract void copyMessage(Message message, String targetFolder) throws IOException;

    /**
     * Move message to target folder
     *
     * @param message      Exchange message
     * @param targetFolder target folder
     * @throws IOException on error
     */
    public abstract void moveMessage(Message message, String targetFolder) throws IOException;

    /**
     * Move folder to target name.
     *
     * @param folderName current folder name/path
     * @param targetName target folder name/path
     * @throws IOException on error
     */
    public abstract void moveFolder(String folderName, String targetName) throws IOException;

    /**
     * Move item from source path to target path.
     *
     * @param sourcePath item source path
     * @param targetPath item target path
     * @throws IOException on error
     */
    public abstract void moveItem(String sourcePath, String targetPath) throws IOException;

    public abstract void moveToTrash(Message message) throws IOException;

    /**
     * Convert keyword value to IMAP flag.
     *
     * @param value keyword value
     * @return IMAP flag
     */
    public String convertKeywordToFlag(String value) {
        // first test for keyword in settings
        Properties flagSettings = Settings.getSubProperties("davmail.imapFlags");
        Enumeration flagSettingsEnum = flagSettings.propertyNames();
        while (flagSettingsEnum.hasMoreElements()) {
            String key = (String) flagSettingsEnum.nextElement();
            if (value.equalsIgnoreCase(flagSettings.getProperty(key))) {
                return key;
            }
        }

        ResourceBundle flagBundle = ResourceBundle.getBundle("imapflags");
        Enumeration<String> flagBundleEnum = flagBundle.getKeys();
        while (flagBundleEnum.hasMoreElements()) {
            String key = flagBundleEnum.nextElement();
            if (value.equalsIgnoreCase(flagBundle.getString(key))) {
                return key;
            }
        }

        // fall back to raw value
        return value;
    }

    /**
     * Convert IMAP flag to keyword value.
     *
     * @param value IMAP flag
     * @return keyword value
     */
    public String convertFlagToKeyword(String value) {
        // first test for flag in settings
        Properties flagSettings = Settings.getSubProperties("davmail.imapFlags");
        String flagValue = flagSettings.getProperty(value);
        if (flagValue != null) {
            return flagValue;
        }

        // fall back to predefined flags
        ResourceBundle flagBundle = ResourceBundle.getBundle("imapflags");
        try {
            return flagBundle.getString(value);
        } catch (MissingResourceException e) {
            // ignore
        }

        // fall back to raw value
        return value;
    }

    protected abstract Set<String> getItemProperties();

    /**
     * Search contacts in provided folder.
     *
     * @param folderPath Exchange folder path
     * @return list of contacts
     * @throws IOException on error
     */
    public List<Contact> getAllContacts(String folderPath) throws IOException {
        return searchContacts(folderPath, ExchangeSession.CONTACT_ATTRIBUTES, isEqualTo("outlookmessageclass", "IPM.Contact"), 0);
    }


    /**
     * Search contacts in provided folder matching the search query.
     *
     * @param folderPath Exchange folder path
     * @param attributes requested attributes
     * @param condition  Exchange search query
     * @param maxCount   maximum item count
     * @return list of contacts
     * @throws IOException on error
     */
    public abstract List<Contact> searchContacts(String folderPath, Set<String> attributes, Condition condition, int maxCount) throws IOException;

    /**
     * Search calendar messages in provided folder.
     *
     * @param folderPath Exchange folder path
     * @return list of calendar messages as Event objects
     * @throws IOException on error
     */
    public abstract List<Event> getEventMessages(String folderPath) throws IOException;

    /**
     * Search calendar events in provided folder.
     *
     * @param folderPath Exchange folder path
     * @return list of calendar events
     * @throws IOException on error
     */
    public List<Event> getAllEvents(String folderPath) throws IOException {
        List<Event> results = searchEvents(folderPath, getCalendarItemCondition(getPastDelayCondition("dtstart")));

        if (!Settings.getBooleanProperty("davmail.caldavDisableTasks", false) && isMainCalendar(folderPath)) {
            // retrieve tasks from main tasks folder
            results.addAll(searchTasksOnly(TASKS));
        }

        return results;
    }

    protected abstract Condition getCalendarItemCondition(Condition dateCondition);

    protected Condition getPastDelayCondition(String attribute) {
        int caldavPastDelay = Settings.getIntProperty("davmail.caldavPastDelay");
        Condition dateCondition = null;
        if (caldavPastDelay != 0) {
            Calendar cal = Calendar.getInstance();
            cal.add(Calendar.DAY_OF_MONTH, -caldavPastDelay);
            dateCondition = gt(attribute, formatSearchDate(cal.getTime()));
        }
        return dateCondition;
    }

    protected Condition getRangeCondition(String timeRangeStart, String timeRangeEnd) throws IOException {
        try {
            SimpleDateFormat parser = getZuluDateFormat();
            MultiCondition andCondition = and();
            if (timeRangeStart != null) {
                andCondition.add(gt("dtend", formatSearchDate(parser.parse(timeRangeStart))));
            }
            if (timeRangeEnd != null) {
                andCondition.add(lt("dtstart", formatSearchDate(parser.parse(timeRangeEnd))));
            }
            return andCondition;
        } catch (ParseException e) {
            throw new IOException(e + " " + e.getMessage());
        }
    }

    /**
     * Search events between start and end.
     *
     * @param folderPath     Exchange folder path
     * @param timeRangeStart date range start in zulu format
     * @param timeRangeEnd   date range start in zulu format
     * @return list of calendar events
     * @throws IOException on error
     */
    public List<Event> searchEvents(String folderPath, String timeRangeStart, String timeRangeEnd) throws IOException {
        Condition dateCondition = getRangeCondition(timeRangeStart, timeRangeEnd);
        Condition condition = getCalendarItemCondition(dateCondition);

        return searchEvents(folderPath, condition);
    }

    /**
     * Search events between start and end, exclude tasks.
     *
     * @param folderPath     Exchange folder path
     * @param timeRangeStart date range start in zulu format
     * @param timeRangeEnd   date range start in zulu format
     * @return list of calendar events
     * @throws IOException on error
     */
    public List<Event> searchEventsOnly(String folderPath, String timeRangeStart, String timeRangeEnd) throws IOException {
        Condition dateCondition = getRangeCondition(timeRangeStart, timeRangeEnd);
        return searchEvents(folderPath, getCalendarItemCondition(dateCondition));
    }

    /**
     * Search tasks only (VTODO).
     *
     * @param folderPath Exchange folder path
     * @return list of tasks
     * @throws IOException on error
     */
    public List<Event> searchTasksOnly(String folderPath) throws IOException {
        return searchEvents(folderPath, and(isEqualTo("outlookmessageclass", "IPM.Task"),
                or(isNull("datecompleted"), getPastDelayCondition("datecompleted"))));
    }

    /**
     * Search calendar events in provided folder.
     *
     * @param folderPath Exchange folder path
     * @param filter     search filter
     * @return list of calendar events
     * @throws IOException on error
     */
    public List<Event> searchEvents(String folderPath, Condition filter) throws IOException {

        Condition privateCondition = null;
        if (isSharedFolder(folderPath) && Settings.getBooleanProperty("davmail.excludePrivateEvents", true)) {
            LOGGER.debug("Shared or public calendar: exclude private events");
            privateCondition = isEqualTo("sensitivity", 0);
        }

        return searchEvents(folderPath, getItemProperties(),
                and(filter, privateCondition));
    }

    /**
     * Search calendar events or messages in provided folder matching the search query.
     *
     * @param folderPath Exchange folder path
     * @param attributes requested attributes
     * @param condition  Exchange search query
     * @return list of calendar messages as Event objects
     * @throws IOException on error
     */
    public abstract List<Event> searchEvents(String folderPath, Set<String> attributes, Condition condition) throws IOException;

    /**
     * convert vcf extension to EML.
     *
     * @param itemName item name
     * @return EML item name
     */
    public String convertItemNameToEML(String itemName) {
        if (itemName.endsWith(".vcf")) {
            return itemName.substring(0, itemName.length() - 3) + "EML";
        } else {
            return itemName;
        }
    }

    /**
     * Get item named eventName in folder
     *
     * @param folderPath Exchange folder path
     * @param itemName   event name
     * @return event object
     * @throws IOException on error
     */
    public abstract Item getItem(String folderPath, String itemName) throws IOException;

    /**
     * Retrieve contact photo attached to contact
     *
     * @param contact address book contact
     * @return contact photo
     * @throws IOException on error
     */
    public abstract ContactPhoto getContactPhoto(Contact contact) throws IOException;


    /**
     * Delete event named itemName in folder
     *
     * @param folderPath Exchange folder path
     * @param itemName   item name
     * @throws IOException on error
     */
    public abstract void deleteItem(String folderPath, String itemName) throws IOException;

    /**
     * Mark event processed named eventName in folder
     *
     * @param folderPath Exchange folder path
     * @param itemName   item name
     * @throws IOException on error
     */
    public abstract void processItem(String folderPath, String itemName) throws IOException;

    /**
     * Replace iCal4 (Snow Leopard) principal paths with mailto expression
     *
     * @param value attendee value or ics line
     * @return fixed value
     */
    protected String replaceIcal4Principal(String value) {
        if (value != null && value.contains("/principals/__uuids__/")) {
            return value.replaceAll("/principals/__uuids__/([^/]*)__AT__([^/]*)/", "mailto:$1@$2");
        } else {
            return value;
        }
    }

    /**
     * Build and send the MIME message for the provided ICS event.
     *
     * @param icsBody event in iCalendar format
     * @return HTTP status
     * @throws IOException on error
     */
    public abstract int sendEvent(String icsBody) throws IOException;

    /**
     * Create or update item (event or contact) on the Exchange server
     *
     * @param folderPath Exchange folder path
     * @param itemName   event name
     * @param itemBody   event body in iCalendar format
     * @param etag       previous event etag to detect concurrent updates
     * @param noneMatch  if-none-match header value
     * @return HTTP response event result (status and etag)
     * @throws IOException on error
     */
    public ItemResult createOrUpdateItem(String folderPath, String itemName, String itemBody, String etag, String noneMatch) throws IOException {
        if (itemBody.startsWith("BEGIN:VCALENDAR")) {
            return internalCreateOrUpdateEvent(folderPath, itemName, "urn:content-classes:appointment", itemBody, etag, noneMatch);
        } else if (itemBody.startsWith("BEGIN:VCARD")) {
            return createOrUpdateContact(folderPath, itemName, itemBody, etag, noneMatch);
        } else {
            throw new IOException(BundleMessage.format("EXCEPTION_INVALID_MESSAGE_CONTENT", itemBody));
        }
    }

    static final String[] VCARD_N_PROPERTIES = {"sn", "givenName", "middlename", "personaltitle", "namesuffix"};
    static final String[] VCARD_ADR_HOME_PROPERTIES = {"homepostofficebox", null, "homeStreet", "homeCity", "homeState", "homePostalCode", "homeCountry"};
    static final String[] VCARD_ADR_WORK_PROPERTIES = {"postofficebox", "roomnumber", "street", "l", "st", "postalcode", "co"};
    static final String[] VCARD_ADR_OTHER_PROPERTIES = {"otherpostofficebox", null, "otherstreet", "othercity", "otherstate", "otherpostalcode", "othercountry"};
    static final String[] VCARD_ORG_PROPERTIES = {"o", "department"};

    protected void convertContactProperties(Map<String, String> properties, String[] contactProperties, List<String> values) {
        for (int i = 0; i < values.size() && i < contactProperties.length; i++) {
            if (contactProperties[i] != null) {
                properties.put(contactProperties[i], values.get(i));
            }
        }
    }

    protected ItemResult createOrUpdateContact(String folderPath, String itemName, String itemBody, String etag, String noneMatch) throws IOException {
        // parse VCARD body to build contact property map
        Map<String, String> properties = new HashMap<String, String>();
        properties.put("outlookmessageclass", "IPM.Contact");

        VObject vcard = new VObject(new ICSBufferedReader(new StringReader(itemBody)));
        for (VProperty property : vcard.getProperties()) {
            if ("FN".equals(property.getKey())) {
                properties.put("cn", property.getValue());
                properties.put("subject", property.getValue());
                properties.put("fileas", property.getValue());

            } else if ("N".equals(property.getKey())) {
                convertContactProperties(properties, VCARD_N_PROPERTIES, property.getValues());
            } else if ("NICKNAME".equals(property.getKey())) {
                properties.put("nickname", property.getValue());
            } else if ("TEL".equals(property.getKey())) {
                if (property.hasParam("TYPE", "cell") || property.hasParam("X-GROUP", "cell")) {
                    properties.put("mobile", property.getValue());
                } else if (property.hasParam("TYPE", "work") || property.hasParam("X-GROUP", "work")) {
                    properties.put("telephoneNumber", property.getValue());
                } else if (property.hasParam("TYPE", "home") || property.hasParam("X-GROUP", "home")) {
                    properties.put("homePhone", property.getValue());
                } else if (property.hasParam("TYPE", "fax")) {
                    if (property.hasParam("TYPE", "home")) {
                        properties.put("homefax", property.getValue());
                    } else {
                        properties.put("facsimiletelephonenumber", property.getValue());
                    }
                } else if (property.hasParam("TYPE", "pager")) {
                    properties.put("pager", property.getValue());
                } else if (property.hasParam("TYPE", "car")) {
                    properties.put("othermobile", property.getValue());
                } else {
                    properties.put("otherTelephone", property.getValue());
                }
            } else if ("ADR".equals(property.getKey())) {
                // address
                if (property.hasParam("TYPE", "home")) {
                    convertContactProperties(properties, VCARD_ADR_HOME_PROPERTIES, property.getValues());
                } else if (property.hasParam("TYPE", "work")) {
                    convertContactProperties(properties, VCARD_ADR_WORK_PROPERTIES, property.getValues());
                    // any other type goes to other address
                } else {
                    convertContactProperties(properties, VCARD_ADR_OTHER_PROPERTIES, property.getValues());
                }
            } else if ("EMAIL".equals(property.getKey())) {
                if (property.hasParam("TYPE", "home")) {
                    properties.put("email2", property.getValue());
                    properties.put("smtpemail2", property.getValue());
                } else if (property.hasParam("TYPE", "other")) {
                    properties.put("email3", property.getValue());
                    properties.put("smtpemail3", property.getValue());
                } else {
                    properties.put("email1", property.getValue());
                    properties.put("smtpemail1", property.getValue());
                }
            } else if ("ORG".equals(property.getKey())) {
                convertContactProperties(properties, VCARD_ORG_PROPERTIES, property.getValues());
            } else if ("URL".equals(property.getKey())) {
                if (property.hasParam("TYPE", "work")) {
                    properties.put("businesshomepage", property.getValue());
                } else if (property.hasParam("TYPE", "home")) {
                    properties.put("personalHomePage", property.getValue());
                } else {
                    // default: set personal home page
                    properties.put("personalHomePage", property.getValue());
                }
            } else if ("TITLE".equals(property.getKey())) {
                properties.put("title", property.getValue());
            } else if ("NOTE".equals(property.getKey())) {
                properties.put("description", property.getValue());
            } else if ("CUSTOM1".equals(property.getKey())) {
                properties.put("extensionattribute1", property.getValue());
            } else if ("CUSTOM2".equals(property.getKey())) {
                properties.put("extensionattribute2", property.getValue());
            } else if ("CUSTOM3".equals(property.getKey())) {
                properties.put("extensionattribute3", property.getValue());
            } else if ("CUSTOM4".equals(property.getKey())) {
                properties.put("extensionattribute4", property.getValue());
            } else if ("ROLE".equals(property.getKey())) {
                properties.put("profession", property.getValue());
            } else if ("X-AIM".equals(property.getKey())) {
                properties.put("im", property.getValue());
            } else if ("BDAY".equals(property.getKey())) {
                properties.put("bday", convertBDayToZulu(property.getValue()));
            } else if ("ANNIVERSARY".equals(property.getKey()) || "X-ANNIVERSARY".equals(property.getKey())) {
                properties.put("anniversary", convertBDayToZulu(property.getValue()));
            } else if ("CATEGORIES".equals(property.getKey())) {
                properties.put("keywords", property.getValue());
            } else if ("CLASS".equals(property.getKey())) {
                if ("PUBLIC".equals(property.getValue())) {
                    properties.put("sensitivity", "0");
                    properties.put("private", "false");
                } else {
                    properties.put("sensitivity", "2");
                    properties.put("private", "true");
                }
            } else if ("SEX".equals(property.getKey())) {
                String propertyValue = property.getValue();
                if ("1".equals(propertyValue)) {
                    properties.put("gender", "2");
                } else if ("2".equals(propertyValue)) {
                    properties.put("gender", "1");
                }
            } else if ("FBURL".equals(property.getKey())) {
                properties.put("fburl", property.getValue());
            } else if ("X-ASSISTANT".equals(property.getKey())) {
                properties.put("secretarycn", property.getValue());
            } else if ("X-MANAGER".equals(property.getKey())) {
                properties.put("manager", property.getValue());
            } else if ("X-SPOUSE".equals(property.getKey())) {
                properties.put("spousecn", property.getValue());
            } else if ("PHOTO".equals(property.getKey())) {
                properties.put("photo", property.getValue());
                properties.put("haspicture", "true");
            }
        }
        LOGGER.debug("Create or update contact " + itemName + ": " + properties);
        // reset missing properties to null
        for (String key : CONTACT_ATTRIBUTES) {
            if (!"imapUid".equals(key) && !"etag".equals(key) && !"urlcompname".equals(key)
                    && !"lastmodified".equals(key) && !"sensitivity".equals(key) &&
                    !properties.containsKey(key)) {
                properties.put(key, null);
            }
        }
        return internalCreateOrUpdateContact(folderPath, itemName, properties, etag, noneMatch);
    }

    public String convertZuluDateToBday(String value) {
        String result = null;
        if (value != null && value.length() > 0) {
            try {
                SimpleDateFormat parser = ExchangeSession.getZuluDateFormat();
                Calendar cal = Calendar.getInstance();
                cal.setTime(parser.parse(value));
                cal.add(Calendar.HOUR_OF_DAY, 12);
                result = ExchangeSession.getVcardBdayFormat().format(cal.getTime());
            } catch (ParseException e) {
                LOGGER.warn("Invalid date: " + value);
            }
        }
        return result;
    }

    public String convertBDayToZulu(String value) {
        String result = null;
        if (value != null && value.length() > 0) {
            try {
                SimpleDateFormat parser;
                if (value.length() == 10) {
                    parser = ExchangeSession.getVcardBdayFormat();
                } else if (value.length() == 15) {
                    parser = new SimpleDateFormat("yyyyMMdd'T'HHmmss", Locale.ENGLISH);
                    parser.setTimeZone(GMT_TIMEZONE);
                } else {
                    parser = ExchangeSession.getExchangeZuluDateFormat();
                }
                result = ExchangeSession.getExchangeZuluDateFormatMillisecond().format(parser.parse(value));
            } catch (ParseException e) {
                LOGGER.warn("Invalid date: " + value);
            }
        }

        return result;
    }


    protected abstract ItemResult internalCreateOrUpdateContact(String folderPath, String itemName, Map<String, String> properties, String etag, String noneMatch) throws IOException;

    protected abstract ItemResult internalCreateOrUpdateEvent(String folderPath, String itemName, String contentClass, String icsBody, String etag, String noneMatch) throws IOException;

    /**
     * Get current Exchange alias name from login name
     *
     * @return user name
     */
    public String getAliasFromLogin() {
        // login is email, not alias
        if (this.userName.indexOf('@') >= 0) {
            return null;
        }
        String result = this.userName;
        // remove domain name
        int index = Math.max(result.indexOf('\\'), result.indexOf('/'));
        if (index >= 0) {
            result = result.substring(index + 1);
        }
        return result;
    }

    protected String getEmailSuffixFromHostname() {
        String domain = httpClient.getHostConfiguration().getHost();
        int start = domain.lastIndexOf('.', domain.lastIndexOf('.') - 1);
        if (start >= 0) {
            return '@' + domain.substring(start + 1);
        } else {
            return '@' + domain;
        }
    }

    /**
     * Test if folderPath is inside user mailbox.
     *
     * @param folderPath absolute folder path
     * @return true if folderPath is a public or shared folder
     */
    public abstract boolean isSharedFolder(String folderPath);

    /**
     * Test if folderPath is main calendar.
     *
     * @param folderPath absolute folder path
     * @return true if folderPath is a public or shared folder
     */
    public abstract boolean isMainCalendar(String folderPath);

    static final String MAILBOX_BASE = "/cn=";

    protected void getEmailAndAliasFromOptions() {
        Cookie[] currentCookies = httpClient.getState().getCookies();
        // get user mail URL from html body
        BufferedReader optionsPageReader = null;
        GetMethod optionsMethod = new GetMethod("/owa/?ae=Options&t=About");
        try {
            DavGatewayHttpClientFacade.executeGetMethod(httpClient, optionsMethod, false);
            optionsPageReader = new BufferedReader(new InputStreamReader(optionsMethod.getResponseBodyAsStream()));
            String line;

            // find email and alias
            while ((line = optionsPageReader.readLine()) != null
                    && (line.indexOf('[') == -1
                    || line.indexOf('@') == -1
                    || line.indexOf(']') == -1)
                    && line.toLowerCase().indexOf(MAILBOX_BASE) == -1) {
            }
            if (line != null) {
                int start = line.toLowerCase().lastIndexOf(MAILBOX_BASE) + MAILBOX_BASE.length();
                int end = line.indexOf('<', start);
                alias = line.substring(start, end);
                end = line.lastIndexOf(']');
                start = line.lastIndexOf('[', end) + 1;
                email = line.substring(start, end);
            }
        } catch (IOException e) {
            // restore cookies on error
            httpClient.getState().addCookies(currentCookies);
            LOGGER.error("Error parsing options page at " + optionsMethod.getPath());
        } finally {
            if (optionsPageReader != null) {
                try {
                    optionsPageReader.close();
                } catch (IOException e) {
                    LOGGER.error("Error parsing options page at " + optionsMethod.getPath());
                }
            }
            optionsMethod.releaseConnection();
        }
    }

    /**
     * Get current user email
     *
     * @return user email
     */
    public String getEmail() {
        return email;
    }

    /**
     * Get current user alias
     *
     * @return user email
     */
    public String getAlias() {
        return alias;
    }

    /**
     * Search global address list
     *
     * @param condition           search filter
     * @param returningAttributes returning attributes
     * @param sizeLimit           size limit
     * @return matching contacts from gal
     * @throws IOException on error
     */
    public abstract Map<String, Contact> galFind(Condition condition, Set<String> returningAttributes, int sizeLimit) throws IOException;

    /**
     * Full Contact attribute list
     */
    public static final Set<String> CONTACT_ATTRIBUTES = new HashSet<String>();

    static {
        CONTACT_ATTRIBUTES.add("imapUid");
        CONTACT_ATTRIBUTES.add("etag");
        CONTACT_ATTRIBUTES.add("urlcompname");

        CONTACT_ATTRIBUTES.add("extensionattribute1");
        CONTACT_ATTRIBUTES.add("extensionattribute2");
        CONTACT_ATTRIBUTES.add("extensionattribute3");
        CONTACT_ATTRIBUTES.add("extensionattribute4");
        CONTACT_ATTRIBUTES.add("bday");
        CONTACT_ATTRIBUTES.add("anniversary");
        CONTACT_ATTRIBUTES.add("businesshomepage");
        CONTACT_ATTRIBUTES.add("personalHomePage");
        CONTACT_ATTRIBUTES.add("cn");
        CONTACT_ATTRIBUTES.add("co");
        CONTACT_ATTRIBUTES.add("department");
        CONTACT_ATTRIBUTES.add("smtpemail1");
        CONTACT_ATTRIBUTES.add("smtpemail2");
        CONTACT_ATTRIBUTES.add("smtpemail3");
        CONTACT_ATTRIBUTES.add("facsimiletelephonenumber");
        CONTACT_ATTRIBUTES.add("givenName");
        CONTACT_ATTRIBUTES.add("homeCity");
        CONTACT_ATTRIBUTES.add("homeCountry");
        CONTACT_ATTRIBUTES.add("homePhone");
        CONTACT_ATTRIBUTES.add("homePostalCode");
        CONTACT_ATTRIBUTES.add("homeState");
        CONTACT_ATTRIBUTES.add("homeStreet");
        CONTACT_ATTRIBUTES.add("homepostofficebox");
        CONTACT_ATTRIBUTES.add("l");
        CONTACT_ATTRIBUTES.add("manager");
        CONTACT_ATTRIBUTES.add("mobile");
        CONTACT_ATTRIBUTES.add("namesuffix");
        CONTACT_ATTRIBUTES.add("nickname");
        CONTACT_ATTRIBUTES.add("o");
        CONTACT_ATTRIBUTES.add("pager");
        CONTACT_ATTRIBUTES.add("personaltitle");
        CONTACT_ATTRIBUTES.add("postalcode");
        CONTACT_ATTRIBUTES.add("postofficebox");
        CONTACT_ATTRIBUTES.add("profession");
        CONTACT_ATTRIBUTES.add("roomnumber");
        CONTACT_ATTRIBUTES.add("secretarycn");
        CONTACT_ATTRIBUTES.add("sn");
        CONTACT_ATTRIBUTES.add("spousecn");
        CONTACT_ATTRIBUTES.add("st");
        CONTACT_ATTRIBUTES.add("street");
        CONTACT_ATTRIBUTES.add("telephoneNumber");
        CONTACT_ATTRIBUTES.add("title");
        CONTACT_ATTRIBUTES.add("description");
        CONTACT_ATTRIBUTES.add("im");
        CONTACT_ATTRIBUTES.add("middlename");
        CONTACT_ATTRIBUTES.add("lastmodified");
        CONTACT_ATTRIBUTES.add("otherstreet");
        CONTACT_ATTRIBUTES.add("otherstate");
        CONTACT_ATTRIBUTES.add("otherpostofficebox");
        CONTACT_ATTRIBUTES.add("otherpostalcode");
        CONTACT_ATTRIBUTES.add("othercountry");
        CONTACT_ATTRIBUTES.add("othercity");
        CONTACT_ATTRIBUTES.add("haspicture");
        CONTACT_ATTRIBUTES.add("keywords");
        CONTACT_ATTRIBUTES.add("othermobile");
        CONTACT_ATTRIBUTES.add("otherTelephone");
        CONTACT_ATTRIBUTES.add("gender");
        CONTACT_ATTRIBUTES.add("private");
        CONTACT_ATTRIBUTES.add("sensitivity");
        CONTACT_ATTRIBUTES.add("fburl");
    }

    /**
     * Get freebusy data string from Exchange.
     *
     * @param attendee attendee email address
     * @param start    start date in Exchange zulu format
     * @param end      end date in Exchange zulu format
     * @param interval freebusy interval in minutes
     * @return freebusy data or null
     * @throws IOException on error
     */
    protected abstract String getFreeBusyData(String attendee, String start, String end, int interval) throws IOException;

    /**
     * Get freebusy info for attendee between start and end date.
     *
     * @param attendee       attendee email
     * @param startDateValue start date
     * @param endDateValue   end date
     * @return FreeBusy info
     * @throws IOException on error
     */
    public FreeBusy getFreebusy(String attendee, String startDateValue, String endDateValue) throws IOException {
        // replace ical encoded attendee name
        attendee = replaceIcal4Principal(attendee);

        // then check that email address is valid to avoid InvalidSmtpAddress error
        if (attendee == null || attendee.indexOf('@') < 0 || attendee.charAt(attendee.length() - 1) == '@') {
            return null;
        }

        if (attendee.startsWith("mailto:") || attendee.startsWith("MAILTO:")) {
            attendee = attendee.substring("mailto:".length());
        }

        SimpleDateFormat exchangeZuluDateFormat = getExchangeZuluDateFormat();
        SimpleDateFormat icalDateFormat = getZuluDateFormat();

        Date startDate;
        Date endDate;
        try {
            if (startDateValue.length() == 8) {
                startDate = parseDate(startDateValue);
            } else {
                startDate = icalDateFormat.parse(startDateValue);
            }
            if (endDateValue.length() == 8) {
                endDate = parseDate(endDateValue);
            } else {
                endDate = icalDateFormat.parse(endDateValue);
            }
        } catch (ParseException e) {
            throw new DavMailException("EXCEPTION_INVALID_DATES", e.getMessage());
        }

        FreeBusy freeBusy = null;
        String fbdata = getFreeBusyData(attendee, exchangeZuluDateFormat.format(startDate), exchangeZuluDateFormat.format(endDate), FreeBusy.INTERVAL);
        if (fbdata != null) {
            freeBusy = new FreeBusy(icalDateFormat, startDate, fbdata);
        }

        if (freeBusy != null && freeBusy.knownAttendee) {
            return freeBusy;
        } else {
            return null;
        }
    }

    protected VObject vTimezone;

    /**
     * Load and return current user OWA timezone.
     *
     * @return current timezone
     */
    public VObject getVTimezone() {
        if (vTimezone == null) {
            // need to load Timezone info from OWA
            loadVtimezone();
        }
        return vTimezone;
    }

    protected abstract void loadVtimezone();

    /**
     * Return internal HttpClient instance
     *
     * @return http client
     */
    public HttpClient getHttpClient() {
        return httpClient;
    }
}
