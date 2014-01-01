/*
 * DavMail POP/IMAP/SMTP/CalDav/LDAP Exchange Gateway
 * Copyright (C) 2010  Mickael Guessant
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
package davmail.exchange.dav;

import davmail.BundleMessage;
import davmail.Settings;
import davmail.exception.*;
import davmail.exchange.*;
import davmail.exchange.condition.*;
import davmail.exchange.entity.*;
import davmail.http.DavGatewayHttpClientFacade;
import davmail.ui.tray.DavGatewayTray;
import davmail.util.IOUtil;
import davmail.util.StringUtil;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.httpclient.*;
import org.apache.commons.httpclient.methods.*;
import org.apache.commons.httpclient.util.URIUtil;
import org.apache.jackrabbit.webdav.DavException;
import org.apache.jackrabbit.webdav.MultiStatus;
import org.apache.jackrabbit.webdav.MultiStatusResponse;
import org.apache.jackrabbit.webdav.client.methods.CopyMethod;
import org.apache.jackrabbit.webdav.client.methods.MoveMethod;
import org.apache.jackrabbit.webdav.client.methods.PropFindMethod;
import org.apache.jackrabbit.webdav.client.methods.PropPatchMethod;
import org.apache.jackrabbit.webdav.property.DavProperty;
import org.apache.jackrabbit.webdav.property.DavPropertyNameSet;
import org.apache.jackrabbit.webdav.property.DavPropertySet;
import org.apache.jackrabbit.webdav.property.PropEntry;
import org.w3c.dom.Node;

import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimePart;
import javax.mail.util.SharedByteArrayInputStream;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.*;
import java.net.NoRouteToHostException;
import java.net.SocketException;
import java.net.URL;
import java.net.UnknownHostException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.GZIPInputStream;

/**
 * Webdav Exchange adapter.
 * Compatible with Exchange 2003 and 2007 with webdav available.
 */
public class DavExchangeSession extends ExchangeSession {
    protected static enum FolderQueryTraversal {
        Shallow, Deep
    }

    protected static final DavPropertyNameSet WELL_KNOWN_FOLDERS = new DavPropertyNameSet();

    static {
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("inbox"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("deleteditems"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("sentitems"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("sendmsg"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("drafts"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("calendar"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("tasks"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("contacts"));
        WELL_KNOWN_FOLDERS.add(Field.getPropertyName("outbox"));
    }

    static final Map<String, String> vTodoToTaskStatusMap = new HashMap<String, String>();
    static final Map<String, String> taskTovTodoStatusMap = new HashMap<String, String>();

    static {
        //taskTovTodoStatusMap.put("0", null);
        taskTovTodoStatusMap.put("1", "IN-PROCESS");
        taskTovTodoStatusMap.put("2", "COMPLETED");
        taskTovTodoStatusMap.put("3", "NEEDS-ACTION");
        taskTovTodoStatusMap.put("4", "CANCELLED");

        //vTodoToTaskStatusMap.put(null, "0");
        vTodoToTaskStatusMap.put("IN-PROCESS", "1");
        vTodoToTaskStatusMap.put("COMPLETED", "2");
        vTodoToTaskStatusMap.put("NEEDS-ACTION", "3");
        vTodoToTaskStatusMap.put("CANCELLED", "4");
    }

    /**
     * Various standard mail boxes Urls
     */
    protected String inboxUrl;
    protected String deleteditemsUrl;
    protected String sentitemsUrl;
    protected String sendmsgUrl;
    protected String draftsUrl;
    protected String calendarUrl;
    protected String tasksUrl;
    protected String contactsUrl;
    protected String outboxUrl;

    protected String inboxName;
    protected String deleteditemsName;
    protected String sentitemsName;
    protected String sendmsgName;
    protected String draftsName;
    protected String calendarName;
    protected String tasksName;
    protected String contactsName;
    protected String outboxName;

    protected static final String USERS = "/users/";

    @Override
    public boolean isExpired() throws NoRouteToHostException, UnknownHostException {
        // experimental: try to reset session timeout
        if (serverVersion.isExchange2007()) {
            GetMethod getMethod = null;
            try {
                getMethod = new GetMethod("/owa/");
                getMethod.setFollowRedirects(false);
                httpClient.executeMethod(getMethod);
            } catch (IOException e) {
                LOGGER.warn(e.getMessage());
            } finally {
                if (getMethod != null) {
                    getMethod.releaseConnection();
                }
            }
        }

        return super.isExpired();
    }


    /**
     * Convert logical or relative folder path to exchange folder path.
     *
     * @param folderPath folder name
     * @return folder path
     */
    public String getFolderPath(String folderPath) {
        String exchangeFolderPath;
        // IMAP path
        if (folderPath.startsWith(INBOX)) {
            exchangeFolderPath = mailPath + inboxName + folderPath.substring(INBOX.length());
        } else if (folderPath.startsWith(TRASH)) {
            exchangeFolderPath = mailPath + deleteditemsName + folderPath.substring(TRASH.length());
        } else if (folderPath.startsWith(DRAFTS)) {
            exchangeFolderPath = mailPath + draftsName + folderPath.substring(DRAFTS.length());
        } else if (folderPath.startsWith(SENT)) {
            exchangeFolderPath = mailPath + sentitemsName + folderPath.substring(SENT.length());
        } else if (folderPath.startsWith(SENDMSG)) {
            exchangeFolderPath = mailPath + sendmsgName + folderPath.substring(SENDMSG.length());
        } else if (folderPath.startsWith(CONTACTS)) {
            exchangeFolderPath = mailPath + contactsName + folderPath.substring(CONTACTS.length());
        } else if (folderPath.startsWith(CALENDAR)) {
            exchangeFolderPath = mailPath + calendarName + folderPath.substring(CALENDAR.length());
        } else if (folderPath.startsWith(TASKS)) {
            exchangeFolderPath = mailPath + tasksName + folderPath.substring(TASKS.length());
        } else if (folderPath.startsWith("public")) {
            exchangeFolderPath = publicFolderUrl + folderPath.substring("public".length());

            // caldav path
        } else if (folderPath.startsWith(USERS)) {
            // get requested principal
            String principal;
            String localPath;
            int principalIndex = folderPath.indexOf('/', USERS.length());
            if (principalIndex >= 0) {
                principal = folderPath.substring(USERS.length(), principalIndex);
                localPath = folderPath.substring(USERS.length() + principal.length() + 1);
                if (localPath.startsWith(LOWER_CASE_INBOX) || localPath.startsWith(INBOX)) {
                    localPath = inboxName + localPath.substring(LOWER_CASE_INBOX.length());
                } else if (localPath.startsWith(CALENDAR)) {
                    localPath = calendarName + localPath.substring(CALENDAR.length());
                } else if (localPath.startsWith(TASKS)) {
                    localPath = tasksName + localPath.substring(TASKS.length());
                } else if (localPath.startsWith(CONTACTS)) {
                    localPath = contactsName + localPath.substring(CONTACTS.length());
                } else if (localPath.startsWith(ADDRESSBOOK)) {
                    localPath = contactsName + localPath.substring(ADDRESSBOOK.length());
                }
            } else {
                principal = folderPath.substring(USERS.length());
                localPath = "";
            }
            if (principal.length() == 0) {
                exchangeFolderPath = rootPath;
            } else if (alias.equalsIgnoreCase(principal) || email.equalsIgnoreCase(principal)) {
                exchangeFolderPath = mailPath + localPath;
            } else {
                LOGGER.debug("Detected shared path for principal " + principal + ", user principal is " + email);
                exchangeFolderPath = rootPath + principal + '/' + localPath;
            }

            // absolute folder path
        } else if (folderPath.startsWith("/")) {
            exchangeFolderPath = folderPath;
        } else {
            exchangeFolderPath = mailPath + folderPath;
        }
        return exchangeFolderPath;
    }

    /**
     * Test if folderPath is inside user mailbox.
     *
     * @param folderPath absolute folder path
     * @return true if folderPath is a public or shared folder
     */
    @Override
    public boolean isSharedFolder(String folderPath) {
        return !getFolderPath(folderPath).toLowerCase().startsWith(mailPath.toLowerCase());
    }

    /**
     * Test if folderPath is main calendar.
     *
     * @param folderPath absolute folder path
     * @return true if folderPath is a public or shared folder
     */
    @Override
    public boolean isMainCalendar(String folderPath) {
        return getFolderPath(folderPath).equalsIgnoreCase(getFolderPath("calendar"));
    }

    /**
     * Build base path for cmd commands (galfind, gallookup).
     *
     * @return cmd base path
     */
    public String getCmdBasePath() {
        if ((serverVersion.isExchange2003() || PUBLIC_ROOT.equals(publicFolderUrl)) && mailPath != null) {
            // public folder is not available => try to use mailbox path
            // Note: This does not work with freebusy, which requires /public/
            return mailPath;
        } else {
            // use public folder url
            return publicFolderUrl;
        }
    }

    /**
     * LDAP to Exchange Criteria Map
     */
    static final HashMap<String, String> GALFIND_CRITERIA_MAP = new HashMap<String, String>();

    static {
        GALFIND_CRITERIA_MAP.put("imapUid", "AN");
        GALFIND_CRITERIA_MAP.put("smtpemail1", "EM");
        GALFIND_CRITERIA_MAP.put("cn", "DN");
        GALFIND_CRITERIA_MAP.put("givenName", "FN");
        GALFIND_CRITERIA_MAP.put("sn", "LN");
        GALFIND_CRITERIA_MAP.put("title", "TL");
        GALFIND_CRITERIA_MAP.put("o", "CP");
        GALFIND_CRITERIA_MAP.put("l", "OF");
        GALFIND_CRITERIA_MAP.put("department", "DP");
    }

    static final HashSet<String> GALLOOKUP_ATTRIBUTES = new HashSet<String>();

    static {
        GALLOOKUP_ATTRIBUTES.add("givenName");
        GALLOOKUP_ATTRIBUTES.add("initials");
        GALLOOKUP_ATTRIBUTES.add("sn");
        GALLOOKUP_ATTRIBUTES.add("street");
        GALLOOKUP_ATTRIBUTES.add("st");
        GALLOOKUP_ATTRIBUTES.add("postalcode");
        GALLOOKUP_ATTRIBUTES.add("co");
        GALLOOKUP_ATTRIBUTES.add("departement");
        GALLOOKUP_ATTRIBUTES.add("mobile");
    }

    /**
     * Exchange to LDAP attribute map
     */
    static final HashMap<String, String> GALFIND_ATTRIBUTE_MAP = new HashMap<String, String>();

    static {
        GALFIND_ATTRIBUTE_MAP.put("uid", "AN");
        GALFIND_ATTRIBUTE_MAP.put("smtpemail1", "EM");
        GALFIND_ATTRIBUTE_MAP.put("cn", "DN");
        GALFIND_ATTRIBUTE_MAP.put("displayName", "DN");
        GALFIND_ATTRIBUTE_MAP.put("telephoneNumber", "PH");
        GALFIND_ATTRIBUTE_MAP.put("l", "OFFICE");
        GALFIND_ATTRIBUTE_MAP.put("o", "CP");
        GALFIND_ATTRIBUTE_MAP.put("title", "TL");

        GALFIND_ATTRIBUTE_MAP.put("givenName", "first");
        GALFIND_ATTRIBUTE_MAP.put("initials", "initials");
        GALFIND_ATTRIBUTE_MAP.put("sn", "last");
        GALFIND_ATTRIBUTE_MAP.put("street", "street");
        GALFIND_ATTRIBUTE_MAP.put("st", "state");
        GALFIND_ATTRIBUTE_MAP.put("postalcode", "zip");
        GALFIND_ATTRIBUTE_MAP.put("co", "country");
        GALFIND_ATTRIBUTE_MAP.put("department", "department");
        GALFIND_ATTRIBUTE_MAP.put("mobile", "mobile");
        GALFIND_ATTRIBUTE_MAP.put("roomnumber", "office");
    }

    boolean disableGalFind;

    protected Map<String, Map<String, String>> galFind(String query) throws IOException {
        Map<String, Map<String, String>> results;
        String path = getCmdBasePath() + "?Cmd=galfind" + query;
        GetMethod getMethod = new GetMethod(path);
        try {
            DavGatewayHttpClientFacade.executeGetMethod(httpClient, getMethod, true);
            results = XMLStreamUtil.getElementContentsAsMap(getMethod.getResponseBodyAsStream(), "item", "AN");
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(path + ": " + results.size() + " result(s)");
            }
        } catch (IOException e) {
            LOGGER.debug("GET " + path + " failed: " + e + ' ' + e.getMessage());
            disableGalFind = true;
            throw e;
        } finally {
            getMethod.releaseConnection();
        }
        return results;
    }


    @Override
    public Map<String, Contact> galFind(Condition condition, Set<String> returningAttributes, int sizeLimit) throws IOException {
        Map<String, Contact> contacts = new HashMap<String, Contact>();
        if (disableGalFind) {
            // do nothing
        } else if (condition instanceof MultiCondition) {
            List<Condition> conditions = ((MultiCondition) condition).getConditions();
            Operator operator = ((MultiCondition) condition).getOperator();
            if (operator == Operator.Or) {
                for (Condition innerCondition : conditions) {
                    contacts.putAll(galFind(innerCondition, returningAttributes, sizeLimit));
                }
            } else if (operator == Operator.And && !conditions.isEmpty()) {
                Map<String, Contact> innerContacts = galFind(conditions.get(0), returningAttributes, sizeLimit);
                for (Contact contact : innerContacts.values()) {
                    if (condition.isMatch(contact)) {
                        contacts.put(contact.getName().toLowerCase(), contact);
                    }
                }
            }
        } else if (condition instanceof AttributeCondition) {
            String searchAttributeName = ((AttributeCondition) condition).getAttributeName();
            String searchAttribute = GALFIND_CRITERIA_MAP.get(searchAttributeName);
            if (searchAttribute != null) {
                String searchValue = ((AttributeCondition) condition).getValue();
                StringBuilder query = new StringBuilder();
                if ("EM".equals(searchAttribute)) {
                    // mail search, split
                    int atIndex = searchValue.indexOf('@');
                    // remove suffix
                    if (atIndex >= 0) {
                        searchValue = searchValue.substring(0, atIndex);
                    }
                    // split firstname.lastname
                    int dotIndex = searchValue.indexOf('.');
                    if (dotIndex >= 0) {
                        // assume mail starts with firstname
                        query.append("&FN=").append(URIUtil.encodeWithinQuery(searchValue.substring(0, dotIndex)));
                        query.append("&LN=").append(URIUtil.encodeWithinQuery(searchValue.substring(dotIndex + 1)));
                    } else {
                        query.append("&FN=").append(URIUtil.encodeWithinQuery(searchValue));
                    }
                } else {
                    query.append('&').append(searchAttribute).append('=').append(URIUtil.encodeWithinQuery(searchValue));
                }
                Map<String, Map<String, String>> results = galFind(query.toString());
                for (Map<String, String> result : results.values()) {
                    DavContact contact = new DavContact(this);
                    contact.setName(result.get("AN"));
                    contact.put("imapUid", result.get("AN"));
                    buildGalfindContact(contact, result);
                    if (needGalLookup(searchAttributeName, returningAttributes)) {
                        galLookup(contact);
                        // iCal fix to suit both iCal 3 and 4:  move cn to sn, remove cn
                    } else if (returningAttributes.contains("apple-serviceslocator")) {
                        if (contact.get("cn") != null && returningAttributes.contains("sn")) {
                            contact.put("sn", contact.get("cn"));
                            contact.remove("cn");
                        }
                    }
                    if (condition.isMatch(contact)) {
                        contacts.put(contact.getName().toLowerCase(), contact);
                    }
                }
            }

        }
        return contacts;
    }

    protected boolean needGalLookup(String searchAttributeName, Set<String> returningAttributes) {
        // return all attributes => call gallookup
        if (returningAttributes == null || returningAttributes.isEmpty()) {
            return true;
            // iCal search, do not call gallookup
        } else if (returningAttributes.contains("apple-serviceslocator")) {
            return false;
            // Lightning search, no need to gallookup
        } else if ("sn".equals(searchAttributeName)) {
            return returningAttributes.contains("sn");
            // search attribute is gallookup attribute, need to fetch value for isMatch
        } else if (GALLOOKUP_ATTRIBUTES.contains(searchAttributeName)) {
            return true;
        }

        for (String attributeName : GALLOOKUP_ATTRIBUTES) {
            if (returningAttributes.contains(attributeName)) {
                return true;
            }
        }
        return false;
    }

    private boolean disableGalLookup;

    /**
     * Get extended address book information for person with gallookup.
     * Does not work with Exchange 2007
     *
     * @param contact galfind contact
     */
    public void galLookup(DavContact contact) {
        if (!disableGalLookup) {
            LOGGER.debug("galLookup(" + contact.get("smtpemail1") + ')');
            GetMethod getMethod = null;
            try {
                getMethod = new GetMethod(URIUtil.encodePathQuery(getCmdBasePath() + "?Cmd=gallookup&ADDR=" + contact.get("smtpemail1")));
                DavGatewayHttpClientFacade.executeGetMethod(httpClient, getMethod, true);
                Map<String, Map<String, String>> results = XMLStreamUtil.getElementContentsAsMap(getMethod.getResponseBodyAsStream(), "person", "alias");
                // add detailed information
                if (!results.isEmpty()) {
                    Map<String, String> personGalLookupDetails = results.get(contact.get("uid").toLowerCase());
                    if (personGalLookupDetails != null) {
                        buildGalfindContact(contact, personGalLookupDetails);
                    }
                }
            } catch (IOException e) {
                LOGGER.warn("Unable to gallookup person: " + contact + ", disable GalLookup");
                disableGalLookup = true;
            } finally {
                if (getMethod != null) {
                    getMethod.releaseConnection();
                }
            }
        }
    }

    protected void buildGalfindContact(DavContact contact, Map<String, String> response) {
        for (Map.Entry<String, String> entry : GALFIND_ATTRIBUTE_MAP.entrySet()) {
            String attributeValue = response.get(entry.getValue());
            if (attributeValue != null) {
                contact.put(entry.getKey(), attributeValue);
            }
        }
    }

    @Override
    protected String getFreeBusyData(String attendee, String start, String end, int interval) throws IOException {
        String freebusyUrl = publicFolderUrl + "/?cmd=freebusy" +
                "&start=" + start +
                "&end=" + end +
                "&interval=" + interval +
                "&u=SMTP:" + attendee;
        GetMethod getMethod = new GetMethod(freebusyUrl);
        getMethod.setRequestHeader("Content-Type", "text/xml");
        String fbdata = null;
        try {
            DavGatewayHttpClientFacade.executeGetMethod(httpClient, getMethod, true);
            fbdata = StringUtil.getLastToken(getMethod.getResponseBodyAsString(), "<a:fbdata>", "</a:fbdata>");
        } finally {
            getMethod.releaseConnection();
        }
        return fbdata;
    }

    /**
     * @inheritDoc
     */
    public DavExchangeSession(String url, String userName, String password) throws IOException {
        super(url, userName, password);
    }

    @Override
    protected void buildSessionInfo(HttpMethod method) throws DavMailException {
        buildMailPath(method);

        // get base http mailbox http urls
        getWellKnownFolders();
    }

    static final String BASE_HREF = "<base href=\"";

    /**
     * Exchange 2003: get mailPath from welcome page
     *
     * @param method current http method
     * @return mail path from body
     */
    protected String getMailpathFromWelcomePage(HttpMethod method) {
        String welcomePageMailPath = null;
        // get user mail URL from html body (multi frame)
        BufferedReader mainPageReader = null;
        try {
            mainPageReader = new BufferedReader(new InputStreamReader(method.getResponseBodyAsStream()));
            //noinspection StatementWithEmptyBody
            String line;
            while ((line = mainPageReader.readLine()) != null && line.toLowerCase().indexOf(BASE_HREF) == -1) {
            }
            if (line != null) {
                // Exchange 2003
                int start = line.toLowerCase().indexOf(BASE_HREF) + BASE_HREF.length();
                int end = line.indexOf('\"', start);
                String mailBoxBaseHref = line.substring(start, end);
                URL baseURL = new URL(mailBoxBaseHref);
                welcomePageMailPath = URIUtil.decode(baseURL.getPath());
                LOGGER.debug("Base href found in body, mailPath is " + welcomePageMailPath);
            }
        } catch (IOException e) {
            LOGGER.error("Error parsing main page at " + method.getPath(), e);
        } finally {
            if (mainPageReader != null) {
                try {
                    mainPageReader.close();
                } catch (IOException e) {
                    LOGGER.error("Error parsing main page at " + method.getPath());
                }
            }
            method.releaseConnection();
        }
        return welcomePageMailPath;
    }

    protected void buildMailPath(HttpMethod method) throws DavMailAuthenticationException {
        // get mailPath from welcome page on Exchange 2003
        mailPath = getMailpathFromWelcomePage(method);

        //noinspection VariableNotUsedInsideIf
        if (mailPath != null) {
            // Exchange 2003
            serverVersion = ExchangeVersion.Exchange2003;
            fixClientHost(method);
            checkPublicFolder();
            try {
                buildEmail(method.getURI().getHost());
            } catch (URIException uriException) {
                LOGGER.warn(uriException);
            }
        } else {
            // Exchange 2007 : get alias and email from options page
            serverVersion = ExchangeVersion.Exchange2007;

            // Gallookup is an Exchange 2003 only feature
            disableGalLookup = true;
            fixClientHost(method);
            getEmailAndAliasFromOptions();

            checkPublicFolder();

            // failover: try to get email through Webdav and Galfind
            if (alias == null || email == null) {
                try {
                    buildEmail(method.getURI().getHost());
                } catch (URIException uriException) {
                    LOGGER.warn(uriException);
                }
            }

            // build standard mailbox link with email
            mailPath = "/exchange/" + email + '/';
        }

        if (mailPath == null || email == null) {
            throw new DavMailAuthenticationException("EXCEPTION_AUTHENTICATION_FAILED_PASSWORD_EXPIRED");
        }
        LOGGER.debug("Current user email is " + email + ", alias is " + alias + ", mailPath is " + mailPath + " on " + serverVersion);
        rootPath = mailPath.substring(0, mailPath.lastIndexOf('/', mailPath.length() - 2) + 1);
    }

    /**
     * Determine user email through various means.
     *
     * @param hostName Exchange server host name for last failover
     */
    public void buildEmail(String hostName) {
        String mailBoxPath = getMailboxPath();
        // mailPath contains either alias or email
        if (mailBoxPath != null && mailBoxPath.indexOf('@') >= 0) {
            email = mailBoxPath;
            alias = getAliasFromMailboxDisplayName();
            if (alias == null) {
                alias = getAliasFromLogin();
            }
        } else {
            // use mailbox name as alias
            alias = mailBoxPath;
            email = getEmail(alias);
            if (email == null) {
                // failover: try to get email from login name
                alias = getAliasFromLogin();
                email = getEmail(alias);
            }
            // another failover : get alias from mailPath display name
            if (email == null) {
                alias = getAliasFromMailboxDisplayName();
                email = getEmail(alias);
            }
            if (email == null) {
                LOGGER.debug("Unable to get user email with alias " + mailBoxPath
                        + " or " + getAliasFromLogin()
                        + " or " + alias
                );
                // last failover: build email from domain name and mailbox display name
                StringBuilder buffer = new StringBuilder();
                // most reliable alias
                if (mailBoxPath != null) {
                    alias = mailBoxPath;
                } else {
                    alias = getAliasFromLogin();
                }
                buffer.append(alias);
                if (alias.indexOf('@') < 0) {
                    buffer.append('@');
                    int dotIndex = hostName.indexOf('.');
                    if (dotIndex >= 0) {
                        buffer.append(hostName.substring(dotIndex + 1));
                    }
                }
                email = buffer.toString();
            }
        }
    }

    /**
     * Get user alias from mailbox display name over Webdav.
     *
     * @return user alias
     */
    public String getAliasFromMailboxDisplayName() {
        if (mailPath == null) {
            return null;
        }
        String displayName = null;
        try {
            Folder rootFolder = getFolder("");
            if (rootFolder == null) {
                LOGGER.warn(new BundleMessage("EXCEPTION_UNABLE_TO_GET_MAIL_FOLDER", mailPath));
            } else {
                displayName = rootFolder.displayName;
            }
        } catch (IOException e) {
            LOGGER.warn(new BundleMessage("EXCEPTION_UNABLE_TO_GET_MAIL_FOLDER", mailPath));
        }
        return displayName;
    }

    /**
     * Get current Exchange alias name from mailbox name
     *
     * @return user name
     */
    protected String getMailboxPath() {
        if (mailPath == null) {
            return null;
        }
        int index = mailPath.lastIndexOf('/', mailPath.length() - 2);
        if (index >= 0 && mailPath.endsWith("/")) {
            return mailPath.substring(index + 1, mailPath.length() - 1);
        } else {
            LOGGER.warn(new BundleMessage("EXCEPTION_INVALID_MAIL_PATH", mailPath));
            return null;
        }
    }

    /**
     * Get user email from global address list (galfind).
     *
     * @param alias user alias
     * @return user email
     */
    public String getEmail(String alias) {
        String emailResult = null;
        if (alias != null && !disableGalFind) {
            try {
                Map<String, Map<String, String>> results = galFind("&AN=" + URIUtil.encodeWithinQuery(alias));
                Map<String, String> result = results.get(alias.toLowerCase());
                if (result != null) {
                    emailResult = result.get("EM");
                }
            } catch (IOException e) {
                // galfind not available
                disableGalFind = true;
                LOGGER.debug("getEmail(" + alias + ") failed");
            }
        }
        return emailResult;
    }

    protected String getURIPropertyIfExists(DavPropertySet properties, String alias) throws URIException {
        DavProperty property = properties.get(Field.getPropertyName(alias));
        if (property == null) {
            return null;
        } else {
            return URIUtil.decode((String) property.getValue());
        }
    }

    // return last folder name from url

    protected String getFolderName(String url) {
        if (url != null) {
            if (url.endsWith("/")) {
                return url.substring(url.lastIndexOf('/', url.length() - 2) + 1, url.length() - 1);
            } else if (url.indexOf('/') > 0) {
                return url.substring(url.lastIndexOf('/') + 1);
            } else {
                return null;
            }
        } else {
            return null;
        }
    }

    protected void fixClientHost(HttpMethod method) {
        try {
            // update client host, workaround for Exchange 2003 mailbox with an Exchange 2007 frontend
            URI currentUri = method.getURI();
            if (currentUri != null && currentUri.getHost() != null && currentUri.getScheme() != null) {
                httpClient.getHostConfiguration().setHost(currentUri.getHost(), currentUri.getPort(), currentUri.getScheme());
            }
        } catch (URIException e) {
            LOGGER.warn("Unable to update http client host:" + e.getMessage(), e);
        }
    }

    protected void checkPublicFolder() {

        Cookie[] currentCookies = httpClient.getState().getCookies();
        // check public folder access
        try {
            publicFolderUrl = httpClient.getHostConfiguration().getHostURL() + PUBLIC_ROOT;
            DavPropertyNameSet davPropertyNameSet = new DavPropertyNameSet();
            davPropertyNameSet.add(Field.getPropertyName("displayname"));
            PropFindMethod propFindMethod = new PropFindMethod(publicFolderUrl, davPropertyNameSet, 0);
            try {
                DavGatewayHttpClientFacade.executeMethod(httpClient, propFindMethod);
            } catch (IOException e) {
                // workaround for NTLM authentication only on /public
                if (!DavGatewayHttpClientFacade.hasNTLMorNegotiate(httpClient)) {
                    DavGatewayHttpClientFacade.addNTLM(httpClient);
                    DavGatewayHttpClientFacade.executeMethod(httpClient, propFindMethod);
                }
            }
            // update public folder URI
            publicFolderUrl = propFindMethod.getURI().getURI();
        } catch (IOException e) {
            // restore cookies on error
            httpClient.getState().addCookies(currentCookies);
            LOGGER.warn("Public folders not available: " + (e.getMessage() == null ? e : e.getMessage()));
            // default public folder path
            publicFolderUrl = PUBLIC_ROOT;
        }
    }

    protected void getWellKnownFolders() throws DavMailException {
        // Retrieve well known URLs
        MultiStatusResponse[] responses;
        try {
            responses = DavGatewayHttpClientFacade.executePropFindMethod(
                    httpClient, URIUtil.encodePath(mailPath), 0, WELL_KNOWN_FOLDERS);
            if (responses.length == 0) {
                throw new WebdavNotAvailableException("EXCEPTION_UNABLE_TO_GET_MAIL_FOLDER", mailPath);
            }
            DavPropertySet properties = responses[0].getProperties(HttpStatus.SC_OK);
            inboxUrl = getURIPropertyIfExists(properties, "inbox");
            inboxName = getFolderName(inboxUrl);
            deleteditemsUrl = getURIPropertyIfExists(properties, "deleteditems");
            deleteditemsName = getFolderName(deleteditemsUrl);
            sentitemsUrl = getURIPropertyIfExists(properties, "sentitems");
            sentitemsName = getFolderName(sentitemsUrl);
            sendmsgUrl = getURIPropertyIfExists(properties, "sendmsg");
            sendmsgName = getFolderName(sendmsgUrl);
            draftsUrl = getURIPropertyIfExists(properties, "drafts");
            draftsName = getFolderName(draftsUrl);
            calendarUrl = getURIPropertyIfExists(properties, "calendar");
            calendarName = getFolderName(calendarUrl);
            tasksUrl = getURIPropertyIfExists(properties, "tasks");
            tasksName = getFolderName(tasksUrl);
            contactsUrl = getURIPropertyIfExists(properties, "contacts");
            contactsName = getFolderName(contactsUrl);
            outboxUrl = getURIPropertyIfExists(properties, "outbox");
            outboxName = getFolderName(outboxUrl);
            // junk folder not available over webdav

            LOGGER.debug("Inbox URL: " + inboxUrl +
                    " Trash URL: " + deleteditemsUrl +
                    " Sent URL: " + sentitemsUrl +
                    " Send URL: " + sendmsgUrl +
                    " Drafts URL: " + draftsUrl +
                    " Calendar URL: " + calendarUrl +
                    " Tasks URL: " + tasksUrl +
                    " Contacts URL: " + contactsUrl +
                    " Outbox URL: " + outboxUrl +
                    " Public folder URL: " + publicFolderUrl
            );
        } catch (IOException e) {
            LOGGER.error(e.getMessage());
            throw new WebdavNotAvailableException("EXCEPTION_UNABLE_TO_GET_MAIL_FOLDER", mailPath);
        }
    }

    static final Map<Operator, String> OPERATOR_MAP = new HashMap<Operator, String>();

    static {
        OPERATOR_MAP.put(Operator.IsEqualTo, " = ");
        OPERATOR_MAP.put(Operator.IsGreaterThanOrEqualTo, " >= ");
        OPERATOR_MAP.put(Operator.IsGreaterThan, " > ");
        OPERATOR_MAP.put(Operator.IsLessThanOrEqualTo, " <= ");
        OPERATOR_MAP.put(Operator.IsLessThan, " < ");
        OPERATOR_MAP.put(Operator.Like, " like ");
        OPERATOR_MAP.put(Operator.IsNull, " is null");
        OPERATOR_MAP.put(Operator.IsFalse, " = false");
        OPERATOR_MAP.put(Operator.IsTrue, " = true");
        OPERATOR_MAP.put(Operator.StartsWith, " = ");
        OPERATOR_MAP.put(Operator.Contains, " = ");
    }

    @Override
    public davmail.exchange.condition.MultiCondition and(Condition... condition) {
        return new DavMultiCondition(Operator.And, condition);
    }

    @Override
    public davmail.exchange.condition.MultiCondition or(Condition... condition) {
        return new DavMultiCondition(Operator.Or, condition);
    }

    @Override
    public Condition not(Condition condition) {
        if (condition == null) {
            return null;
        } else {
            return new DavNotCondition(condition);
        }
    }

    @Override
    public Condition isEqualTo(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.IsEqualTo, value);
    }

    @Override
    public Condition isEqualTo(String attributeName, int value) {
        return new DavAttributeCondition(attributeName, Operator.IsEqualTo, value);
    }

    @Override
    public Condition headerIsEqualTo(String headerName, String value) {
        return new DavHeaderCondition(headerName, Operator.IsEqualTo, value);
    }

    @Override
    public Condition gte(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.IsGreaterThanOrEqualTo, value);
    }

    @Override
    public Condition lte(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.IsLessThanOrEqualTo, value);
    }

    @Override
    public Condition lt(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.IsLessThan, value);
    }

    @Override
    public Condition gt(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.IsGreaterThan, value);
    }

    @Override
    public Condition contains(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.Like, value);
    }

    @Override
    public Condition startsWith(String attributeName, String value) {
        return new DavAttributeCondition(attributeName, Operator.StartsWith, value);
    }

    @Override
    public Condition isNull(String attributeName) {
        return new DavMonoCondition(attributeName, Operator.IsNull);
    }

    @Override
    public Condition isTrue(String attributeName) {
        if (serverVersion.isExchange2003() && "deleted".equals(attributeName)) {
            return isEqualTo(attributeName, "1");
        } else {
            return new DavMonoCondition(attributeName, Operator.IsTrue);
        }
    }

    @Override
    public Condition isFalse(String attributeName) {
        if (serverVersion.isExchange2003() && "deleted".equals(attributeName)) {
            return or(isEqualTo(attributeName, "0"), isNull(attributeName));
        } else {
            return new DavMonoCondition(attributeName, Operator.IsFalse);
        }
    }


    protected Folder buildFolder(MultiStatusResponse entity) throws IOException {
        String href = URIUtil.decode(entity.getHref());
        Folder folder = new Folder(this);
        DavPropertySet properties = entity.getProperties(HttpStatus.SC_OK);
        folder.displayName = getPropertyIfExists(properties, "displayname");
        folder.folderClass = getPropertyIfExists(properties, "folderclass");
        folder.hasChildren = "1".equals(getPropertyIfExists(properties, "hassubs"));
        folder.noInferiors = "1".equals(getPropertyIfExists(properties, "nosubs"));
        folder.count = getIntPropertyIfExists(properties, "count");
        folder.unreadCount = getIntPropertyIfExists(properties, "unreadcount");
        // fake recent value
        folder.recent = folder.unreadCount;
        folder.ctag = getPropertyIfExists(properties, "contenttag");
        folder.etag = getPropertyIfExists(properties, "lastmodified");

        folder.uidNext = getIntPropertyIfExists(properties, "uidNext");

        // replace well known folder names
        if (inboxUrl != null && href.startsWith(inboxUrl)) {
            folder.folderPath = href.replaceFirst(inboxUrl, INBOX);
        } else if (sentitemsUrl != null && href.startsWith(sentitemsUrl)) {
            folder.folderPath = href.replaceFirst(sentitemsUrl, SENT);
        } else if (draftsUrl != null && href.startsWith(draftsUrl)) {
            folder.folderPath = href.replaceFirst(draftsUrl, DRAFTS);
        } else if (deleteditemsUrl != null && href.startsWith(deleteditemsUrl)) {
            folder.folderPath = href.replaceFirst(deleteditemsUrl, TRASH);
        } else if (calendarUrl != null && href.startsWith(calendarUrl)) {
            folder.folderPath = href.replaceFirst(calendarUrl, CALENDAR);
        } else if (contactsUrl != null && href.startsWith(contactsUrl)) {
            folder.folderPath = href.replaceFirst(contactsUrl, CONTACTS);
        } else {
            int index = href.indexOf(mailPath.substring(0, mailPath.length() - 1));
            if (index >= 0) {
                if (index + mailPath.length() > href.length()) {
                    folder.folderPath = "";
                } else {
                    folder.folderPath = href.substring(index + mailPath.length());
                }
            } else {
                try {
                    URI folderURI = new URI(href, false);
                    folder.folderPath = folderURI.getPath();
                } catch (URIException e) {
                    throw new DavMailException("EXCEPTION_INVALID_FOLDER_URL", href);
                }
            }
        }
        if (folder.folderPath.endsWith("/")) {
            folder.folderPath = folder.folderPath.substring(0, folder.folderPath.length() - 1);
        }
        return folder;
    }

    protected static final Set<String> FOLDER_PROPERTIES = new HashSet<String>();

    static {
        FOLDER_PROPERTIES.add("displayname");
        FOLDER_PROPERTIES.add("folderclass");
        FOLDER_PROPERTIES.add("hassubs");
        FOLDER_PROPERTIES.add("nosubs");
        FOLDER_PROPERTIES.add("count");
        FOLDER_PROPERTIES.add("unreadcount");
        FOLDER_PROPERTIES.add("contenttag");
        FOLDER_PROPERTIES.add("lastmodified");
        FOLDER_PROPERTIES.add("uidNext");
    }

    protected static final DavPropertyNameSet FOLDER_PROPERTIES_NAME_SET = new DavPropertyNameSet();

    static {
        for (String attribute : FOLDER_PROPERTIES) {
            FOLDER_PROPERTIES_NAME_SET.add(Field.getPropertyName(attribute));
        }
    }

    /**
     * @inheritDoc
     */
    @Override
    protected Folder internalGetFolder(String folderPath) throws IOException {
        MultiStatusResponse[] responses = DavGatewayHttpClientFacade.executePropFindMethod(
                httpClient, URIUtil.encodePath(getFolderPath(folderPath)), 0, FOLDER_PROPERTIES_NAME_SET);
        Folder folder = null;
        if (responses.length > 0) {
            folder = buildFolder(responses[0]);
            folder.folderPath = folderPath;
        }
        return folder;
    }

    /**
     * @inheritDoc
     */
    @Override
    public List<Folder> getSubFolders(String folderPath, Condition condition, boolean recursive) throws IOException {
        boolean isPublic = folderPath.startsWith("/public");
        FolderQueryTraversal mode = (!isPublic && recursive) ? FolderQueryTraversal.Deep : FolderQueryTraversal.Shallow;
        List<Folder> folders = new ArrayList<Folder>();

        MultiStatusResponse[] responses = searchItems(folderPath, FOLDER_PROPERTIES, and(isTrue("isfolder"), isFalse("ishidden"), condition), mode, 0);

        for (MultiStatusResponse response : responses) {
            Folder folder = buildFolder(response);
            folders.add(buildFolder(response));
            if (isPublic && recursive) {
                getSubFolders(folder.folderPath, condition, recursive);
            }
        }
        return folders;
    }

    /**
     * @inheritDoc
     */
    @Override
    public int createFolder(String folderPath, String folderClass, Map<String, String> properties) throws IOException {
        Set<PropertyValue> propertyValues = new HashSet<PropertyValue>();
        if (properties != null) {
            for (Map.Entry<String, String> entry : properties.entrySet()) {
                propertyValues.add(Field.createPropertyValue(entry.getKey(), entry.getValue()));
            }
        }
        propertyValues.add(Field.createPropertyValue("folderclass", folderClass));

        // standard MkColMethod does not take properties, override PropPatchMethod instead
        ExchangePropPatchMethod method = new ExchangePropPatchMethod(URIUtil.encodePath(getFolderPath(folderPath)), propertyValues) {
            @Override
            public String getName() {
                return "MKCOL";
            }
        };
        int status = DavGatewayHttpClientFacade.executeHttpMethod(httpClient, method);
        if (status == HttpStatus.SC_MULTI_STATUS) {
            status = method.getResponseStatusCode();
        }
        return status;
    }

    /**
     * @inheritDoc
     */
    @Override
    public int updateFolder(String folderPath, Map<String, String> properties) throws IOException {
        Set<PropertyValue> propertyValues = new HashSet<PropertyValue>();
        if (properties != null) {
            for (Map.Entry<String, String> entry : properties.entrySet()) {
                propertyValues.add(Field.createPropertyValue(entry.getKey(), entry.getValue()));
            }
        }

        // standard MkColMethod does not take properties, override PropPatchMethod instead
        ExchangePropPatchMethod method = new ExchangePropPatchMethod(URIUtil.encodePath(getFolderPath(folderPath)), propertyValues);
        int status = DavGatewayHttpClientFacade.executeHttpMethod(httpClient, method);
        if (status == HttpStatus.SC_MULTI_STATUS) {
            status = method.getResponseStatusCode();
        }
        return status;
    }

    /**
     * @inheritDoc
     */
    @Override
    public void deleteFolder(String folderPath) throws IOException {
        DavGatewayHttpClientFacade.executeDeleteMethod(httpClient, URIUtil.encodePath(getFolderPath(folderPath)));
    }

    /**
     * @inheritDoc
     */
    @Override
    public void moveFolder(String folderPath, String targetPath) throws IOException {
        MoveMethod method = new MoveMethod(URIUtil.encodePath(getFolderPath(folderPath)),
                URIUtil.encodePath(getFolderPath(targetPath)), false);
        try {
            int statusCode = httpClient.executeMethod(method);
            if (statusCode == HttpStatus.SC_PRECONDITION_FAILED) {
                throw new HttpPreconditionFailedException(BundleMessage.format("EXCEPTION_UNABLE_TO_MOVE_FOLDER"));
            } else if (statusCode != HttpStatus.SC_CREATED) {
                throw DavGatewayHttpClientFacade.buildHttpException(method);
            } else if (folderPath.equalsIgnoreCase("/users/" + getEmail() + "/calendar")) {
                // calendar renamed, need to reload well known folders 
                getWellKnownFolders();
            }
        } finally {
            method.releaseConnection();
        }
    }

    /**
     * @inheritDoc
     */
    @Override
    public void moveItem(String sourcePath, String targetPath) throws IOException {
        MoveMethod method = new MoveMethod(URIUtil.encodePath(getFolderPath(sourcePath)),
                URIUtil.encodePath(getFolderPath(targetPath)), false);
        moveItem(method);
    }

    protected void moveItem(MoveMethod method) throws IOException {
        try {
            int statusCode = httpClient.executeMethod(method);
            if (statusCode == HttpStatus.SC_PRECONDITION_FAILED) {
                throw new DavMailException("EXCEPTION_UNABLE_TO_MOVE_ITEM");
            } else if (statusCode != HttpStatus.SC_CREATED && statusCode != HttpStatus.SC_OK) {
                throw DavGatewayHttpClientFacade.buildHttpException(method);
            }
        } finally {
            method.releaseConnection();
        }
    }

    protected String getPropertyIfExists(DavPropertySet properties, String alias) {
        DavProperty property = properties.get(Field.getResponsePropertyName(alias));
        if (property == null) {
            return null;
        } else {
            Object value = property.getValue();
            if (value instanceof Node) {
                return ((Node) value).getTextContent();
            } else if (value instanceof List) {
                StringBuilder buffer = new StringBuilder();
                for (Object node : (List) value) {
                    if (buffer.length() > 0) {
                        buffer.append(',');
                    }
                    if (node instanceof Node) {
                        // jackrabbit
                        buffer.append(((Node) node).getTextContent());
                    } else {
                        // ExchangeDavMethod
                        buffer.append(node);
                    }
                }
                return buffer.toString();
            } else {
                return (String) value;
            }
        }
    }

    protected String getURLPropertyIfExists(DavPropertySet properties, String alias) throws URIException {
        String result = getPropertyIfExists(properties, alias);
        if (result != null) {
            result = URIUtil.decode(result);
        }
        return result;
    }

    protected int getIntPropertyIfExists(DavPropertySet properties, String alias) {
        DavProperty property = properties.get(Field.getPropertyName(alias));
        if (property == null) {
            return 0;
        } else {
            return Integer.parseInt((String) property.getValue());
        }
    }

    protected long getLongPropertyIfExists(DavPropertySet properties, String alias) {
        DavProperty property = properties.get(Field.getPropertyName(alias));
        if (property == null) {
            return 0;
        } else {
            return Long.parseLong((String) property.getValue());
        }
    }

    protected double getDoublePropertyIfExists(DavPropertySet properties, String alias) {
        DavProperty property = properties.get(Field.getResponsePropertyName(alias));
        if (property == null) {
            return 0;
        } else {
            return Double.parseDouble((String) property.getValue());
        }
    }

    protected byte[] getBinaryPropertyIfExists(DavPropertySet properties, String alias) {
        byte[] property = null;
        String base64Property = getPropertyIfExists(properties, alias);
        if (base64Property != null) {
            try {
                property = Base64.decodeBase64(base64Property.getBytes("ASCII"));
            } catch (UnsupportedEncodingException e) {
                LOGGER.warn(e);
            }
        }
        return property;
    }


    protected DavMessage buildMessage(MultiStatusResponse responseEntity) throws URIException, DavMailException {
        DavMessage message = new DavMessage(this);

        message.messageUrl = URIUtil.decode(responseEntity.getHref());
        DavPropertySet properties = responseEntity.getProperties(HttpStatus.SC_OK);

        message.permanentUrl = getURLPropertyIfExists(properties, "permanenturl");
        message.size = getIntPropertyIfExists(properties, "messageSize");
        message.uid = getPropertyIfExists(properties, "uid");
        message.contentClass = getPropertyIfExists(properties, "contentclass");
        message.imapUid = getLongPropertyIfExists(properties, "imapUid");
        message.read = "1".equals(getPropertyIfExists(properties, "read"));
        message.junk = "1".equals(getPropertyIfExists(properties, "junk"));
        message.flagged = "2".equals(getPropertyIfExists(properties, "flagStatus"));
        message.draft = (getIntPropertyIfExists(properties, "messageFlags") & 8) != 0;
        String lastVerbExecuted = getPropertyIfExists(properties, "lastVerbExecuted");
        message.answered = "102".equals(lastVerbExecuted) || "103".equals(lastVerbExecuted);
        message.forwarded = "104".equals(lastVerbExecuted);
        message.date = convertDateFromExchange(getPropertyIfExists(properties, "date"));
        message.deleted = "1".equals(getPropertyIfExists(properties, "deleted"));

        String lastmodified = convertDateFromExchange(getPropertyIfExists(properties, "lastmodified"));
        message.recent = !message.read && lastmodified != null && lastmodified.equals(message.date);

        message.keywords = getPropertyIfExists(properties, "keywords");

        if (LOGGER.isDebugEnabled()) {
            StringBuilder buffer = new StringBuilder();
            buffer.append("Message");
            if (message.imapUid != 0) {
                buffer.append(" IMAP uid: ").append(message.imapUid);
            }
            if (message.uid != null) {
                buffer.append(" uid: ").append(message.uid);
            }
            buffer.append(" href: ").append(responseEntity.getHref()).append(" permanenturl:").append(message.permanentUrl);
            LOGGER.debug(buffer.toString());
        }
        return message;
    }

    @Override
    public MessageList searchMessages(String folderPath, Set<String> attributes, Condition condition) throws IOException {
        MessageList messages = new MessageList();
        int maxCount = Settings.getIntProperty("davmail.folderSizeLimit", 0);
        MultiStatusResponse[] responses = searchItems(folderPath, attributes, and(isFalse("isfolder"), isFalse("ishidden"), condition), FolderQueryTraversal.Shallow, maxCount);

        for (MultiStatusResponse response : responses) {
            DavMessage message = buildMessage(response);
            message.messageList = messages;
            messages.add(message);
        }
        Collections.sort(messages);
        return messages;
    }

    /**
     * @inheritDoc
     */
    @Override
    public List<Contact> searchContacts(String folderPath, Set<String> attributes, Condition condition, int maxCount) throws IOException {
        List<Contact> contacts = new ArrayList<Contact>();
        MultiStatusResponse[] responses = searchItems(folderPath, attributes,
                and(isEqualTo("outlookmessageclass", "IPM.Contact"), isFalse("isfolder"), isFalse("ishidden"), condition),
                FolderQueryTraversal.Shallow, maxCount);
        for (MultiStatusResponse response : responses) {
            contacts.add(new DavContact(this, response));
        }
        return contacts;
    }

    /**
     * Common item properties
     */
    protected static final Set<String> ITEM_PROPERTIES = new HashSet<String>();

    static {
        ITEM_PROPERTIES.add("etag");
        ITEM_PROPERTIES.add("displayname");
        // calendar CdoInstanceType
        ITEM_PROPERTIES.add("instancetype");
        ITEM_PROPERTIES.add("urlcompname");
        ITEM_PROPERTIES.add("subject");
        ITEM_PROPERTIES.add("contentclass");
    }

    protected Set<String> getItemProperties() {
        return ITEM_PROPERTIES;
    }


    /**
     * @inheritDoc
     */
    @Override
    public List<Event> getEventMessages(String folderPath) throws IOException {
        return searchEvents(folderPath, ITEM_PROPERTIES,
                and(isEqualTo("contentclass", "urn:content-classes:calendarmessage"),
                        or(isNull("processed"), isFalse("processed"))));
    }


    @Override
    public List<Event> searchEvents(String folderPath, Set<String> attributes, Condition condition) throws IOException {
        MultiStatusResponse[] responses = searchItems(folderPath, attributes, and(isFalse("isfolder"), isFalse("ishidden"), condition), FolderQueryTraversal.Shallow, 0);
        List<Event> events = new ArrayList<Event>(responses.length);
        for (MultiStatusResponse response : responses) {
            String instancetype = getPropertyIfExists(response.getProperties(HttpStatus.SC_OK), "instancetype");
            DavEvent event = new DavEvent(this, response);
            //noinspection VariableNotUsedInsideIf
            if (instancetype == null) {
                // check ics content
                try {
                    event.getBody();
                    // getBody success => add event or task
                    events.add(event);
                } catch (IOException e) {
                    // invalid event: exclude from list
                    LOGGER.warn("Invalid event " + event.displayName + " found at " + response.getHref(), e);
                }
            } else {
                events.add(event);
            }
        }
        return events;
    }

    @Override
    protected Condition getCalendarItemCondition(Condition dateCondition) {
        boolean caldavEnableLegacyTasks = Settings.getBooleanProperty("davmail.caldavEnableLegacyTasks", false);
        if (caldavEnableLegacyTasks) {
            // return tasks created in calendar folder
            return or(isNull("instancetype"),
                    isEqualTo("instancetype", 1),
                    and(isEqualTo("instancetype", 0), dateCondition));
        } else {
            // instancetype 0 single appointment / 1 master recurring appointment
            return and(isEqualTo("outlookmessageclass", "IPM.Appointment"),
                    or(isEqualTo("instancetype", 1),
                            and(isEqualTo("instancetype", 0), dateCondition)));
        }
    }

    protected MultiStatusResponse[] searchItems(String folderPath, Set<String> attributes, Condition condition,
                                                FolderQueryTraversal folderQueryTraversal, int maxCount) throws IOException {
        String folderUrl;
        if (folderPath.startsWith("http")) {
            folderUrl = folderPath;
        } else {
            folderUrl = getFolderPath(folderPath);
        }
        StringBuilder searchRequest = new StringBuilder();
        searchRequest.append("SELECT ")
                .append(Field.getRequestPropertyString("permanenturl"));
        if (attributes != null) {
            for (String attribute : attributes) {
                searchRequest.append(',').append(Field.getRequestPropertyString(attribute));
            }
        }
        searchRequest.append(" FROM SCOPE('").append(folderQueryTraversal).append(" TRAVERSAL OF \"").append(folderUrl).append("\"')");
        if (condition != null) {
            searchRequest.append(" WHERE ");
            condition.appendTo(searchRequest);
        }
        searchRequest.append(" ORDER BY ").append(Field.getRequestPropertyString("imapUid")).append(" DESC");
        DavGatewayTray.debug(new BundleMessage("LOG_SEARCH_QUERY", searchRequest));
        MultiStatusResponse[] responses = DavGatewayHttpClientFacade.executeSearchMethod(
                httpClient, encodeAndFixUrl(folderUrl), searchRequest.toString(), maxCount);
        DavGatewayTray.debug(new BundleMessage("LOG_SEARCH_RESULT", responses.length));
        return responses;
    }

    protected static final Set<String> EVENT_REQUEST_PROPERTIES = new HashSet<String>();

    static {
        EVENT_REQUEST_PROPERTIES.add("permanenturl");
        EVENT_REQUEST_PROPERTIES.add("urlcompname");
        EVENT_REQUEST_PROPERTIES.add("etag");
        EVENT_REQUEST_PROPERTIES.add("contentclass");
        EVENT_REQUEST_PROPERTIES.add("displayname");
        EVENT_REQUEST_PROPERTIES.add("subject");
    }

    protected static final DavPropertyNameSet EVENT_REQUEST_PROPERTIES_NAME_SET = new DavPropertyNameSet();

    static {
        for (String attribute : EVENT_REQUEST_PROPERTIES) {
            EVENT_REQUEST_PROPERTIES_NAME_SET.add(Field.getPropertyName(attribute));
        }

    }

    @Override
    public Item getItem(String folderPath, String itemName) throws IOException {
        String emlItemName = convertItemNameToEML(itemName);
        String itemPath = getFolderPath(folderPath) + '/' + emlItemName;
        MultiStatusResponse[] responses = null;
        try {
            try {
                responses = DavGatewayHttpClientFacade.executePropFindMethod(httpClient, URIUtil.encodePath(itemPath), 0, EVENT_REQUEST_PROPERTIES_NAME_SET);
            } catch (HttpNotFoundException e) {
                // ignore
            }
            if (responses == null || responses.length == 0 && isMainCalendar(folderPath)) {
                if (itemName.endsWith(".ics")) {
                    itemName = itemName.substring(0, itemName.length() - 3) + "EML";
                }
                // look for item in tasks folder
                responses = DavGatewayHttpClientFacade.executePropFindMethod(httpClient, URIUtil.encodePath(getFolderPath(TASKS) + '/' + emlItemName), 0, EVENT_REQUEST_PROPERTIES_NAME_SET);
            }
            if (responses == null || responses.length == 0) {
                throw new HttpNotFoundException(itemPath + " not found");
            }
        } catch (HttpNotFoundException e) {
            try {
                LOGGER.debug(itemPath + " not found, searching by urlcompname");
                // failover: try to get event by displayname
                responses = searchItems(folderPath, EVENT_REQUEST_PROPERTIES, isEqualTo("urlcompname", emlItemName), FolderQueryTraversal.Shallow, 1);
                if (responses.length == 0 && isMainCalendar(folderPath)) {
                    responses = searchItems(TASKS, EVENT_REQUEST_PROPERTIES, isEqualTo("urlcompname", emlItemName), FolderQueryTraversal.Shallow, 1);
                }
                if (responses.length == 0) {
                    throw new HttpNotFoundException(itemPath + " not found");
                }
            } catch (HttpNotFoundException e2) {
                LOGGER.debug("last failover: search all items");
                List<Event> events = getAllEvents(folderPath);
                for (Event event : events) {
                    if (itemName.equals(event.getName())) {
                        responses = DavGatewayHttpClientFacade.executePropFindMethod(httpClient, encodeAndFixUrl(((DavEvent) event).getPermanentUrl()), 0, EVENT_REQUEST_PROPERTIES_NAME_SET);
                        break;
                    }
                }
                if (responses == null || responses.length == 0) {
                    throw new HttpNotFoundException(itemPath + " not found");
                }
                LOGGER.warn("search by urlcompname failed, actual value is " + getPropertyIfExists(responses[0].getProperties(HttpStatus.SC_OK), "urlcompname"));
            }
        }
        // build item
        String contentClass = getPropertyIfExists(responses[0].getProperties(HttpStatus.SC_OK), "contentclass");
        String urlcompname = getPropertyIfExists(responses[0].getProperties(HttpStatus.SC_OK), "urlcompname");
        if ("urn:content-classes:person".equals(contentClass)) {
            // retrieve Contact properties
            List<Contact> contacts = searchContacts(folderPath, CONTACT_ATTRIBUTES,
                    isEqualTo("urlcompname", StringUtil.decodeUrlcompname(urlcompname)), 1);
            if (contacts.isEmpty()) {
                LOGGER.warn("Item found, but unable to build contact");
                throw new HttpNotFoundException(itemPath + " not found");
            }
            return contacts.get(0);
        } else if ("urn:content-classes:appointment".equals(contentClass)
                || "urn:content-classes:calendarmessage".equals(contentClass)
                || "urn:content-classes:task".equals(contentClass)) {
            return new DavEvent(this, responses[0]);
        } else {
            LOGGER.warn("wrong contentclass on item " + itemPath + ": " + contentClass);
            // return item anyway
            return new DavEvent(this, responses[0]);
        }

    }

    @Override
    public ContactPhoto getContactPhoto(Contact contact) throws IOException {
        ContactPhoto contactPhoto = null;
        if ("true".equals(contact.get("haspicture"))) {
            final GetMethod method = new GetMethod(URIUtil.encodePath(contact.getHref()) + "/ContactPicture.jpg");
            method.setRequestHeader("Translate", "f");
            method.setRequestHeader("Accept-Encoding", "gzip");

            InputStream inputStream = null;
            try {
                DavGatewayHttpClientFacade.executeGetMethod(httpClient, method, true);
                if (DavGatewayHttpClientFacade.isGzipEncoded(method)) {
                    inputStream = (new GZIPInputStream(method.getResponseBodyAsStream()));
                } else {
                    inputStream = method.getResponseBodyAsStream();
                }

                contactPhoto = new ContactPhoto();
                contactPhoto.contentType = "image/jpeg";

                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                InputStream partInputStream = inputStream;
                byte[] bytes = new byte[8192];
                int length;
                while ((length = partInputStream.read(bytes)) > 0) {
                    baos.write(bytes, 0, length);
                }
                contactPhoto.content = new String(Base64.encodeBase64(baos.toByteArray()));
            } finally {
                if (inputStream != null) {
                    try {
                        inputStream.close();
                    } catch (IOException e) {
                        LOGGER.debug(e);
                    }
                }
                method.releaseConnection();
            }
        }
        return contactPhoto;
    }

    @Override
    public int sendEvent(String icsBody) throws IOException {
        String itemName = UUID.randomUUID().toString() + ".EML";
        byte[] mimeContent = (new DavEvent(this, getFolderPath(DRAFTS), itemName, "urn:content-classes:calendarmessage", icsBody, null, null)).createMimeContent();
        if (mimeContent == null) {
            // no recipients, cancel
            return HttpStatus.SC_NO_CONTENT;
        } else {
            sendMessage(mimeContent);
            return HttpStatus.SC_OK;
        }
    }

    @Override
    public void deleteItem(String folderPath, String itemName) throws IOException {
        String eventPath = URIUtil.encodePath(getFolderPath(folderPath) + '/' + convertItemNameToEML(itemName));
        int status = DavGatewayHttpClientFacade.executeDeleteMethod(httpClient, eventPath);
        if (status == HttpStatus.SC_NOT_FOUND && isMainCalendar(folderPath)) {
            // retry in tasks folder
            eventPath = URIUtil.encodePath(getFolderPath(TASKS) + '/' + convertItemNameToEML(itemName));
            status = DavGatewayHttpClientFacade.executeDeleteMethod(httpClient, eventPath);
        }
        if (status == HttpStatus.SC_NOT_FOUND) {
            LOGGER.debug("Unable to delete " + itemName + ": item not found");
        }
    }

    @Override
    public void processItem(String folderPath, String itemName) throws IOException {
        String eventPath = URIUtil.encodePath(getFolderPath(folderPath) + '/' + convertItemNameToEML(itemName));
        // do not delete calendar messages, mark read and processed
        ArrayList<PropEntry> list = new ArrayList<PropEntry>();
        list.add(Field.createDavProperty("processed", "true"));
        list.add(Field.createDavProperty("read", "1"));
        PropPatchMethod patchMethod = new PropPatchMethod(eventPath, list);
        DavGatewayHttpClientFacade.executeMethod(httpClient, patchMethod);
    }

    @Override
    public ItemResult internalCreateOrUpdateEvent(String folderPath, String itemName, String contentClass, String icsBody, String etag, String noneMatch) throws IOException {
        return new DavEvent(this, getFolderPath(folderPath), itemName, contentClass, icsBody, etag, noneMatch).createOrUpdate();
    }

    /**
     * create a fake event to get VTIMEZONE body
     */
    @Override
    protected void loadVtimezone() {
        try {
            // create temporary folder
            String folderPath = getFolderPath("davmailtemp");
            createCalendarFolder(folderPath, null);

            String fakeEventUrl = null;
            if (serverVersion.isExchange2003()) {
                PostMethod postMethod = new PostMethod(URIUtil.encodePath(folderPath));
                postMethod.addParameter("Cmd", "saveappt");
                postMethod.addParameter("FORMTYPE", "appointment");
                try {
                    // create fake event
                    int statusCode = httpClient.executeMethod(postMethod);
                    if (statusCode == HttpStatus.SC_OK) {
                        fakeEventUrl = StringUtil.getToken(postMethod.getResponseBodyAsString(), "<span id=\"itemHREF\">", "</span>");
                        if (fakeEventUrl != null) {
                            fakeEventUrl = URIUtil.decode(fakeEventUrl);
                        }
                    }
                } finally {
                    postMethod.releaseConnection();
                }
            }
            // failover for Exchange 2007, use PROPPATCH with forced timezone
            if (fakeEventUrl == null) {
                ArrayList<PropEntry> propertyList = new ArrayList<PropEntry>();
                propertyList.add(Field.createDavProperty("contentclass", "urn:content-classes:appointment"));
                propertyList.add(Field.createDavProperty("outlookmessageclass", "IPM.Appointment"));
                propertyList.add(Field.createDavProperty("instancetype", "0"));

                // get forced timezone id from settings
                String timezoneId = Settings.getProperty("davmail.timezoneId");
                if (timezoneId == null) {
                    // get timezoneid from OWA settings
                    timezoneId = getTimezoneIdFromExchange();
                }
                // without a timezoneId, use Exchange timezone
                if (timezoneId != null) {
                    propertyList.add(Field.createDavProperty("timezoneid", timezoneId));
                }
                String patchMethodUrl = folderPath + '/' + UUID.randomUUID().toString() + ".EML";
                PropPatchMethod patchMethod = new PropPatchMethod(URIUtil.encodePath(patchMethodUrl), propertyList);
                try {
                    int statusCode = httpClient.executeMethod(patchMethod);
                    if (statusCode == HttpStatus.SC_MULTI_STATUS) {
                        fakeEventUrl = patchMethodUrl;
                    }
                } finally {
                    patchMethod.releaseConnection();
                }
            }
            if (fakeEventUrl != null) {
                // get fake event body
                GetMethod getMethod = new GetMethod(URIUtil.encodePath(fakeEventUrl));
                getMethod.setRequestHeader("Translate", "f");
                try {
                    httpClient.executeMethod(getMethod);
                    this.vTimezone = new VObject("BEGIN:VTIMEZONE" +
                            StringUtil.getToken(getMethod.getResponseBodyAsString(), "BEGIN:VTIMEZONE", "END:VTIMEZONE") +
                            "END:VTIMEZONE\r\n");
                } finally {
                    getMethod.releaseConnection();
                }
            }

            // delete temporary folder
            deleteFolder("davmailtemp");
        } catch (IOException e) {
            LOGGER.warn("Unable to get VTIMEZONE info: " + e, e);
        }
    }

    protected String getTimezoneIdFromExchange() {
        String timezoneId = null;
        String timezoneName = null;
        try {
            Set<String> attributes = new HashSet<String>();
            attributes.add("roamingdictionary");

            MultiStatusResponse[] responses = searchItems("/users/" + getEmail() + "/NON_IPM_SUBTREE", attributes, isEqualTo("messageclass", "IPM.Configuration.OWA.UserOptions"), DavExchangeSession.FolderQueryTraversal.Deep, 1);
            if (responses.length == 1) {
                byte[] roamingdictionary = getBinaryPropertyIfExists(responses[0].getProperties(HttpStatus.SC_OK), "roamingdictionary");
                if (roamingdictionary != null) {
                    timezoneName = getTimezoneNameFromRoamingDictionary(roamingdictionary);
                    if (timezoneName != null) {
                        timezoneId = ResourceBundle.getBundle("timezoneids").getString(timezoneName);
                    }
                }
            }
        } catch (MissingResourceException e) {
            LOGGER.warn("Unable to retrieve Exchange timezone id for name " + timezoneName);
        } catch (UnsupportedEncodingException e) {
            LOGGER.warn("Unable to retrieve Exchange timezone id: " + e.getMessage(), e);
        } catch (IOException e) {
            LOGGER.warn("Unable to retrieve Exchange timezone id: " + e.getMessage(), e);
        }
        return timezoneId;
    }

    protected String getTimezoneNameFromRoamingDictionary(byte[] roamingdictionary) {
        String timezoneName = null;
        XMLStreamReader reader;
        try {
            reader = XMLStreamUtil.createXMLStreamReader(roamingdictionary);
            while (reader.hasNext()) {
                reader.next();
                if (XMLStreamUtil.isStartTag(reader, "e")
                        && "18-timezone".equals(reader.getAttributeValue(null, "k"))) {
                    String value = reader.getAttributeValue(null, "v");
                    if (value != null && value.startsWith("18-")) {
                        timezoneName = value.substring(3);
                    }
                }
            }

        } catch (XMLStreamException e) {
            LOGGER.error("Error while parsing RoamingDictionary: " + e, e);
        }
        return timezoneName;
    }

    @Override
    protected ItemResult internalCreateOrUpdateContact(String folderPath, String itemName, Map<String, String> properties, String etag, String noneMatch) throws IOException {
        return new DavContact(this, getFolderPath(folderPath), itemName, properties, etag, noneMatch).createOrUpdate();
    }

    protected List<PropEntry> buildProperties(Map<String, String> properties) {
        ArrayList<PropEntry> list = new ArrayList<PropEntry>();
        if (properties != null) {
            for (Map.Entry<String, String> entry : properties.entrySet()) {
                if ("read".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("read", entry.getValue()));
                } else if ("junk".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("junk", entry.getValue()));
                } else if ("flagged".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("flagStatus", entry.getValue()));
                } else if ("answered".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("lastVerbExecuted", entry.getValue()));
                    if ("102".equals(entry.getValue())) {
                        list.add(Field.createDavProperty("iconIndex", "261"));
                    }
                } else if ("forwarded".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("lastVerbExecuted", entry.getValue()));
                    if ("104".equals(entry.getValue())) {
                        list.add(Field.createDavProperty("iconIndex", "262"));
                    }
                } else if ("bcc".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("bcc", entry.getValue()));
                } else if ("deleted".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("deleted", entry.getValue()));
                } else if ("datereceived".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("datereceived", entry.getValue()));
                } else if ("keywords".equals(entry.getKey())) {
                    list.add(Field.createDavProperty("keywords", entry.getValue()));
                }
            }
        }
        return list;
    }

    /**
     * Create message in specified folder.
     * Will overwrite an existing message with same messageName in the same folder
     *
     * @param folderPath  Exchange folder path
     * @param messageName message name
     * @param properties  message properties (flags)
     * @param mimeMessage MIME message
     * @throws IOException when unable to create message
     */
    @Override
    public void createMessage(String folderPath, String messageName, HashMap<String, String> properties, MimeMessage mimeMessage) throws IOException {
        String messageUrl = URIUtil.encodePathQuery(getFolderPath(folderPath) + '/' + messageName);
        PropPatchMethod patchMethod;
        List<PropEntry> davProperties = buildProperties(properties);

        if (properties != null && properties.containsKey("draft")) {
            // note: draft is readonly after create, create the message first with requested messageFlags
            davProperties.add(Field.createDavProperty("messageFlags", properties.get("draft")));
        }
        if (properties != null && properties.containsKey("mailOverrideFormat")) {
            davProperties.add(Field.createDavProperty("mailOverrideFormat", properties.get("mailOverrideFormat")));
        }
        if (properties != null && properties.containsKey("messageFormat")) {
            davProperties.add(Field.createDavProperty("messageFormat", properties.get("messageFormat")));
        }
        if (!davProperties.isEmpty()) {
            patchMethod = new PropPatchMethod(messageUrl, davProperties);
            try {
                // update message with blind carbon copy and other flags
                int statusCode = httpClient.executeMethod(patchMethod);
                if (statusCode != HttpStatus.SC_MULTI_STATUS) {
                    throw new DavMailException("EXCEPTION_UNABLE_TO_CREATE_MESSAGE", messageUrl, statusCode, ' ', patchMethod.getStatusLine());
                }

            } finally {
                patchMethod.releaseConnection();
            }
        }

        // update message body
        PutMethod putmethod = new PutMethod(messageUrl);
        putmethod.setRequestHeader("Translate", "f");
        putmethod.setRequestHeader("Content-Type", "message/rfc822");

        try {
            // use same encoding as client socket reader
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            mimeMessage.writeTo(baos);
            baos.close();
            putmethod.setRequestEntity(new ByteArrayRequestEntity(baos.toByteArray()));
            int code = httpClient.executeMethod(putmethod);

            // workaround for misconfigured Exchange server
            if (code == HttpStatus.SC_NOT_ACCEPTABLE) {
                LOGGER.warn("Draft message creation failed, failover to property update. Note: attachments are lost");

                ArrayList<PropEntry> propertyList = new ArrayList<PropEntry>();
                propertyList.add(Field.createDavProperty("to", mimeMessage.getHeader("to", ",")));
                propertyList.add(Field.createDavProperty("cc", mimeMessage.getHeader("cc", ",")));
                propertyList.add(Field.createDavProperty("message-id", mimeMessage.getHeader("message-id", ",")));

                MimePart mimePart = mimeMessage;
                if (mimeMessage.getContent() instanceof MimeMultipart) {
                    MimeMultipart multiPart = (MimeMultipart) mimeMessage.getContent();
                    for (int i = 0; i < multiPart.getCount(); i++) {
                        String contentType = multiPart.getBodyPart(i).getContentType();
                        if (contentType.startsWith("text/")) {
                            mimePart = (MimePart) multiPart.getBodyPart(i);
                            break;
                        }
                    }
                }

                String contentType = mimePart.getContentType();

                if (contentType.startsWith("text/plain")) {
                    propertyList.add(Field.createDavProperty("description", (String) mimePart.getContent()));
                } else if (contentType.startsWith("text/html")) {
                    propertyList.add(Field.createDavProperty("htmldescription", (String) mimePart.getContent()));
                } else {
                    LOGGER.warn("Unsupported content type: " + contentType + " message body will be empty");
                }

                propertyList.add(Field.createDavProperty("subject", mimeMessage.getHeader("subject", ",")));
                PropPatchMethod propPatchMethod = new PropPatchMethod(messageUrl, propertyList);
                try {
                    int patchStatus = DavGatewayHttpClientFacade.executeHttpMethod(httpClient, propPatchMethod);
                    if (patchStatus == HttpStatus.SC_MULTI_STATUS) {
                        code = HttpStatus.SC_OK;
                    }
                } finally {
                    propPatchMethod.releaseConnection();
                }
            }

            if (code != HttpStatus.SC_OK && code != HttpStatus.SC_CREATED) {

                // first delete draft message
                if (!davProperties.isEmpty()) {
                    try {
                        DavGatewayHttpClientFacade.executeDeleteMethod(httpClient, messageUrl);
                    } catch (IOException e) {
                        LOGGER.warn("Unable to delete draft message");
                    }
                }
                if (code == HttpStatus.SC_INSUFFICIENT_STORAGE) {
                    throw new InsufficientStorageException(putmethod.getStatusText());
                } else {
                    throw new DavMailException("EXCEPTION_UNABLE_TO_CREATE_MESSAGE", messageUrl, code, ' ', putmethod.getStatusLine());
                }
            }
        } catch (MessagingException e) {
            throw new IOException(e.getMessage());
        } finally {
            putmethod.releaseConnection();
        }

        try {
            // need to update bcc after put
            if (mimeMessage.getHeader("Bcc") != null) {
                davProperties = new ArrayList<PropEntry>();
                davProperties.add(Field.createDavProperty("bcc", mimeMessage.getHeader("Bcc", ",")));
                patchMethod = new PropPatchMethod(messageUrl, davProperties);
                try {
                    // update message with blind carbon copy
                    int statusCode = httpClient.executeMethod(patchMethod);
                    if (statusCode != HttpStatus.SC_MULTI_STATUS) {
                        throw new DavMailException("EXCEPTION_UNABLE_TO_CREATE_MESSAGE", messageUrl, statusCode, ' ', patchMethod.getStatusLine());
                    }

                } finally {
                    patchMethod.releaseConnection();
                }
            }
        } catch (MessagingException e) {
            throw new IOException(e.getMessage());
        }

    }

    /**
     * @inheritDoc
     */
    @Override
    public void updateMessage(Message message, Map<String, String> properties) throws IOException {
        PropPatchMethod patchMethod = new PropPatchMethod(encodeAndFixUrl(message.permanentUrl), buildProperties(properties)) {
            @Override
            protected void processResponseBody(HttpState httpState, HttpConnection httpConnection) {
                // ignore response body, sometimes invalid with exchange mapi properties
            }
        };
        try {
            int statusCode = httpClient.executeMethod(patchMethod);
            if (statusCode != HttpStatus.SC_MULTI_STATUS) {
                throw new DavMailException("EXCEPTION_UNABLE_TO_UPDATE_MESSAGE");
            }

        } finally {
            patchMethod.releaseConnection();
        }
    }

    /**
     * @inheritDoc
     */
    @Override
    public void deleteMessage(Message message) throws IOException {
        LOGGER.debug("Delete " + message.permanentUrl + " (" + message.messageUrl + ')');
        DavGatewayHttpClientFacade.executeDeleteMethod(httpClient, encodeAndFixUrl(message.permanentUrl));
    }

    /**
     * Send message.
     *
     * @param messageBody MIME message body
     * @throws IOException on error
     */
    public void sendMessage(byte[] messageBody) throws IOException {
        try {
            sendMessage(new MimeMessage(null, new SharedByteArrayInputStream(messageBody)));
        } catch (MessagingException e) {
            throw new IOException(e.getMessage());
        }
    }

    //protected static final long MAPI_SEND_NO_RICH_INFO = 0x00010000L;
    protected static final long ENCODING_PREFERENCE = 0x00020000L;
    protected static final long ENCODING_MIME = 0x00040000L;
    //protected static final long BODY_ENCODING_HTML = 0x00080000L;
    protected static final long BODY_ENCODING_TEXT_AND_HTML = 0x00100000L;
    //protected static final long MAC_ATTACH_ENCODING_UUENCODE = 0x00200000L;
    //protected static final long MAC_ATTACH_ENCODING_APPLESINGLE = 0x00400000L;
    //protected static final long MAC_ATTACH_ENCODING_APPLEDOUBLE = 0x00600000L;
    //protected static final long OOP_DONT_LOOKUP = 0x10000000L;

    @Override
    public void sendMessage(MimeMessage mimeMessage) throws IOException {
        try {
            // need to create draft first
            String itemName = UUID.randomUUID().toString() + ".EML";
            HashMap<String, String> properties = new HashMap<String, String>();
            properties.put("draft", "9");
            String contentType = mimeMessage.getContentType();
            if (contentType != null && contentType.startsWith("text/plain")) {
                properties.put("messageFormat", "1");
            } else {
                properties.put("mailOverrideFormat", String.valueOf(ENCODING_PREFERENCE | ENCODING_MIME | BODY_ENCODING_TEXT_AND_HTML));
                properties.put("messageFormat", "2");
            }
            createMessage(DRAFTS, itemName, properties, mimeMessage);
            MoveMethod method = new MoveMethod(URIUtil.encodePath(getFolderPath(DRAFTS + '/' + itemName)),
                    URIUtil.encodePath(getFolderPath(SENDMSG)), false);
            // set header if saveInSent is disabled 
            if (!Settings.getBooleanProperty("davmail.smtpSaveInSent", true)) {
                method.setRequestHeader("Saveinsent", "f");
            }
            moveItem(method);
        } catch (MessagingException e) {
            throw new IOException(e.getMessage());
        }
    }

    // wrong hostname fix flag
    protected boolean restoreHostName;

    /**
     * @inheritDoc
     */
    @Override
    public byte[] getContent(Message message) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        InputStream contentInputStream;
        try {
            try {
                try {
                    contentInputStream = getContentInputStream(message.messageUrl);
                } catch (UnknownHostException e) {
                    // failover for misconfigured Exchange server, replace host name in url
                    restoreHostName = true;
                    contentInputStream = getContentInputStream(message.messageUrl);
                }
            } catch (HttpNotFoundException e) {
                LOGGER.debug("Message not found at: " + message.messageUrl + ", retrying with permanenturl");
                contentInputStream = getContentInputStream(message.permanentUrl);
            }

            try {
                IOUtil.write(contentInputStream, baos);
            } finally {
                contentInputStream.close();
            }

        } catch (LoginTimeoutException e) {
            // throw error on expired session
            LOGGER.warn(e.getMessage());
            throw e;
        } catch (SocketException e) {
            // throw error on broken connection
            LOGGER.warn(e.getMessage());
            throw e;
        } catch (IOException e) {
            LOGGER.warn("Broken message at: " + message.messageUrl + " permanentUrl: " + message.permanentUrl + ", trying to rebuild from properties");

            try {
                DavPropertyNameSet messageProperties = new DavPropertyNameSet();
                messageProperties.add(Field.getPropertyName("contentclass"));
                messageProperties.add(Field.getPropertyName("message-id"));
                messageProperties.add(Field.getPropertyName("from"));
                messageProperties.add(Field.getPropertyName("to"));
                messageProperties.add(Field.getPropertyName("cc"));
                messageProperties.add(Field.getPropertyName("subject"));
                messageProperties.add(Field.getPropertyName("date"));
                messageProperties.add(Field.getPropertyName("htmldescription"));
                messageProperties.add(Field.getPropertyName("body"));
                PropFindMethod propFindMethod = new PropFindMethod(encodeAndFixUrl(message.permanentUrl), messageProperties, 0);
                DavGatewayHttpClientFacade.executeMethod(httpClient, propFindMethod);
                MultiStatus responses = propFindMethod.getResponseBodyAsMultiStatus();
                if (responses.getResponses().length > 0) {
                    MimeMessage mimeMessage = new MimeMessage((Session) null);

                    DavPropertySet properties = responses.getResponses()[0].getProperties(HttpStatus.SC_OK);
                    String propertyValue = getPropertyIfExists(properties, "contentclass");
                    if (propertyValue != null) {
                        mimeMessage.addHeader("Content-class", propertyValue);
                    }
                    propertyValue = getPropertyIfExists(properties, "date");
                    if (propertyValue != null) {
                        mimeMessage.setSentDate(parseDateFromExchange(propertyValue));
                    }
                    propertyValue = getPropertyIfExists(properties, "from");
                    if (propertyValue != null) {
                        mimeMessage.addHeader("From", propertyValue);
                    }
                    propertyValue = getPropertyIfExists(properties, "to");
                    if (propertyValue != null) {
                        mimeMessage.addHeader("To", propertyValue);
                    }
                    propertyValue = getPropertyIfExists(properties, "cc");
                    if (propertyValue != null) {
                        mimeMessage.addHeader("Cc", propertyValue);
                    }
                    propertyValue = getPropertyIfExists(properties, "subject");
                    if (propertyValue != null) {
                        mimeMessage.setSubject(propertyValue);
                    }
                    propertyValue = getPropertyIfExists(properties, "htmldescription");
                    if (propertyValue != null) {
                        mimeMessage.setContent(propertyValue, "text/html; charset=UTF-8");
                    } else {
                        propertyValue = getPropertyIfExists(properties, "body");
                        if (propertyValue != null) {
                            mimeMessage.setText(propertyValue);
                        }
                    }
                    mimeMessage.writeTo(baos);
                }
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug("Rebuilt message content: " + new String(baos.toByteArray()));
                }
            } catch (IOException e2) {
                LOGGER.warn(e2);
            } catch (DavException e2) {
                LOGGER.warn(e2);
            } catch (MessagingException e2) {
                LOGGER.warn(e2);
            }
            // other exception
            if (baos.size() == 0 && Settings.getBooleanProperty("davmail.deleteBroken")) {
                LOGGER.warn("Deleting broken message at: " + message.messageUrl + " permanentUrl: " + message.permanentUrl);
                try {
                    message.delete();
                } catch (IOException ioe) {
                    LOGGER.warn("Unable to delete broken message at: " + message.permanentUrl);
                }
                throw e;
            }
        }

        return baos.toByteArray();
    }

    protected String getEscapedUrlFromPath(String escapedPath) throws URIException {
        URI uri = new URI(httpClient.getHostConfiguration().getHostURL(), true);
        uri.setEscapedPath(escapedPath);
        return uri.getEscapedURI();
    }

    public String encodeAndFixUrl(String url) throws URIException {
        String originalUrl = URIUtil.encodePath(url);
        if (restoreHostName && originalUrl.startsWith("http")) {
            String targetPath = new URI(originalUrl, true).getEscapedPath();
            originalUrl = getEscapedUrlFromPath(targetPath);
        }
        return originalUrl;
    }

    protected InputStream getContentInputStream(String url) throws IOException {
        String encodedUrl = encodeAndFixUrl(url);

        final GetMethod method = new GetMethod(encodedUrl);
        method.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
        method.setRequestHeader("Translate", "f");
        method.setRequestHeader("Accept-Encoding", "gzip");

        InputStream inputStream;
        try {
            DavGatewayHttpClientFacade.executeGetMethod(httpClient, method, true);
            if (DavGatewayHttpClientFacade.isGzipEncoded(method)) {
                inputStream = new GZIPInputStream(method.getResponseBodyAsStream());
            } else {
                inputStream = method.getResponseBodyAsStream();
            }
            inputStream = new FilterInputStream(inputStream) {
                int totalCount;
                int lastLogCount;

                @Override
                public int read(byte[] buffer, int offset, int length) throws IOException {
                    int count = super.read(buffer, offset, length);
                    totalCount += count;
                    if (totalCount - lastLogCount > 1024 * 128) {
                        DavGatewayTray.debug(new BundleMessage("LOG_DOWNLOAD_PROGRESS", String.valueOf(totalCount / 1024), method.getURI()));
                        DavGatewayTray.switchIcon();
                        lastLogCount = totalCount;
                    }
                    return count;
                }

                @Override
                public void close() throws IOException {
                    try {
                        super.close();
                    } finally {
                        method.releaseConnection();
                    }
                }
            };

        } catch (HttpException e) {
            method.releaseConnection();
            LOGGER.warn("Unable to retrieve message at: " + url);
            throw e;
        }
        return inputStream;
    }

    /**
     * @inheritDoc
     */
    @Override
    public void moveMessage(Message message, String targetFolder) throws IOException {
        try {
            moveMessage(message.permanentUrl, targetFolder);
        } catch (HttpNotFoundException e) {
            LOGGER.debug("404 not found at permanenturl: " + message.permanentUrl + ", retry with messageurl");
            moveMessage(message.messageUrl, targetFolder);
        }
    }

    protected void moveMessage(String sourceUrl, String targetFolder) throws IOException {
        String targetPath = URIUtil.encodePath(getFolderPath(targetFolder)) + '/' + UUID.randomUUID().toString();
        MoveMethod method = new MoveMethod(URIUtil.encodePath(sourceUrl), targetPath, false);
        // allow rename if a message with the same name exists
        method.addRequestHeader("Allow-Rename", "t");
        try {
            int statusCode = httpClient.executeMethod(method);
            if (statusCode == HttpStatus.SC_PRECONDITION_FAILED) {
                throw new DavMailException("EXCEPTION_UNABLE_TO_MOVE_MESSAGE");
            } else if (statusCode != HttpStatus.SC_CREATED) {
                throw DavGatewayHttpClientFacade.buildHttpException(method);
            }
        } finally {
            method.releaseConnection();
        }
    }

    /**
     * @inheritDoc
     */
    @Override
    public void copyMessage(Message message, String targetFolder) throws IOException {
        try {
            copyMessage(message.permanentUrl, targetFolder);
        } catch (HttpNotFoundException e) {
            LOGGER.debug("404 not found at permanenturl: " + message.permanentUrl + ", retry with messageurl");
            copyMessage(message.messageUrl, targetFolder);
        }
    }

    protected void copyMessage(String sourceUrl, String targetFolder) throws IOException {
        String targetPath = URIUtil.encodePath(getFolderPath(targetFolder)) + '/' + UUID.randomUUID().toString();
        CopyMethod method = new CopyMethod(URIUtil.encodePath(sourceUrl), targetPath, false);
        // allow rename if a message with the same name exists
        method.addRequestHeader("Allow-Rename", "t");
        try {
            int statusCode = httpClient.executeMethod(method);
            if (statusCode == HttpStatus.SC_PRECONDITION_FAILED) {
                throw new DavMailException("EXCEPTION_UNABLE_TO_COPY_MESSAGE");
            } else if (statusCode != HttpStatus.SC_CREATED) {
                throw DavGatewayHttpClientFacade.buildHttpException(method);
            }
        } finally {
            method.releaseConnection();
        }
    }

    @Override
    public void moveToTrash(Message message) throws IOException {
        String destination = URIUtil.encodePath(deleteditemsUrl) + '/' + UUID.randomUUID().toString();
        LOGGER.debug("Deleting : " + message.permanentUrl + " to " + destination);
        MoveMethod method = new MoveMethod(encodeAndFixUrl(message.permanentUrl), destination, false);
        method.addRequestHeader("Allow-rename", "t");

        int status = DavGatewayHttpClientFacade.executeHttpMethod(httpClient, method);
        // do not throw error if already deleted
        if (status != HttpStatus.SC_CREATED && status != HttpStatus.SC_NOT_FOUND) {
            throw DavGatewayHttpClientFacade.buildHttpException(method);
        }
        if (method.getResponseHeader("Location") != null) {
            destination = method.getResponseHeader("Location").getValue();
        }

        LOGGER.debug("Deleted to :" + destination);
    }

    protected String getItemProperty(String permanentUrl, String propertyName) throws IOException, DavException {
        String result = null;
        DavPropertyNameSet davPropertyNameSet = new DavPropertyNameSet();
        davPropertyNameSet.add(Field.getPropertyName(propertyName));
        PropFindMethod propFindMethod = new PropFindMethod(encodeAndFixUrl(permanentUrl), davPropertyNameSet, 0);
        try {
            try {
                DavGatewayHttpClientFacade.executeHttpMethod(httpClient, propFindMethod);
            } catch (UnknownHostException e) {
                propFindMethod.releaseConnection();
                // failover for misconfigured Exchange server, replace host name in url
                restoreHostName = true;
                propFindMethod = new PropFindMethod(encodeAndFixUrl(permanentUrl), davPropertyNameSet, 0);
                DavGatewayHttpClientFacade.executeHttpMethod(httpClient, propFindMethod);
            }

            MultiStatus responses = propFindMethod.getResponseBodyAsMultiStatus();
            if (responses.getResponses().length > 0) {
                DavPropertySet properties = responses.getResponses()[0].getProperties(HttpStatus.SC_OK);
                result = getPropertyIfExists(properties, propertyName);
            }
        } finally {
            propFindMethod.releaseConnection();
        }
        return result;
    }

    protected String convertDateFromExchange(String exchangeDateValue) throws DavMailException {
        String zuluDateValue = null;
        if (exchangeDateValue != null) {
            try {
                zuluDateValue = getZuluDateFormat().format(getExchangeZuluDateFormatMillisecond().parse(exchangeDateValue));
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", exchangeDateValue);
            }
        }
        return zuluDateValue;
    }

    protected static final Map<String, String> importanceToPriorityMap = new HashMap<String, String>();

    static {
        importanceToPriorityMap.put("high", "1");
        importanceToPriorityMap.put("normal", "5");
        importanceToPriorityMap.put("low", "9");
    }

    protected static final Map<String, String> priorityToImportanceMap = new HashMap<String, String>();

    static {
        priorityToImportanceMap.put("1", "high");
        priorityToImportanceMap.put("5", "normal");
        priorityToImportanceMap.put("9", "low");
    }

    protected String convertPriorityFromExchange(String exchangeImportanceValue) {
        String value = null;
        if (exchangeImportanceValue != null) {
            value = importanceToPriorityMap.get(exchangeImportanceValue);
        }
        return value;
    }

    protected String convertPriorityToExchange(String vTodoPriorityValue) {
        String value = null;
        if (vTodoPriorityValue != null) {
            value = priorityToImportanceMap.get(vTodoPriorityValue);
        }
        return value;
    }


    /**
     * Format date to exchange search format.
     *
     * @param date date object
     * @return formatted search date
     */
    @Override
    public String formatSearchDate(Date date) {
        SimpleDateFormat dateFormatter = new SimpleDateFormat(YYYY_MM_DD_HH_MM_SS, Locale.ENGLISH);
        dateFormatter.setTimeZone(GMT_TIMEZONE);
        return dateFormatter.format(date);
    }

    protected String convertTaskDateToZulu(String value) {
        String result = null;
        if (value != null && value.length() > 0) {
            try {
                SimpleDateFormat parser;
                if (value.length() == 8) {
                    parser = new SimpleDateFormat("yyyyMMdd", Locale.ENGLISH);
                    parser.setTimeZone(GMT_TIMEZONE);
                } else if (value.length() == 15) {
                    parser = new SimpleDateFormat("yyyyMMdd'T'HHmmss", Locale.ENGLISH);
                    parser.setTimeZone(GMT_TIMEZONE);
                } else if (value.length() == 16) {
                    parser = new SimpleDateFormat("yyyyMMdd'T'HHmmss'Z'", Locale.ENGLISH);
                    parser.setTimeZone(GMT_TIMEZONE);
                } else {
                    parser = ExchangeSession.getExchangeZuluDateFormat();
                }
                Calendar calendarValue = Calendar.getInstance(GMT_TIMEZONE);
                calendarValue.setTime(parser.parse(value));
                // zulu time: add 12 hours
                if (value.length() == 16) {
                    calendarValue.add(Calendar.HOUR, 12);
                }
                calendarValue.set(Calendar.HOUR, 0);
                calendarValue.set(Calendar.MINUTE, 0);
                calendarValue.set(Calendar.SECOND, 0);
                result = ExchangeSession.getExchangeZuluDateFormatMillisecond().format(calendarValue.getTime());
            } catch (ParseException e) {
                LOGGER.warn("Invalid date: " + value);
            }
        }

        return result;
    }

    protected String convertDateFromExchangeToTaskDate(String exchangeDateValue) throws DavMailException {
        String result = null;
        if (exchangeDateValue != null) {
            try {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd", Locale.ENGLISH);
                dateFormat.setTimeZone(GMT_TIMEZONE);
                result = dateFormat.format(getExchangeZuluDateFormatMillisecond().parse(exchangeDateValue));
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", exchangeDateValue);
            }
        }
        return result;
    }

    protected Date parseDateFromExchange(String exchangeDateValue) throws DavMailException {
        Date result = null;
        if (exchangeDateValue != null) {
            try {
                result = getExchangeZuluDateFormatMillisecond().parse(exchangeDateValue);
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", exchangeDateValue);
            }
        }
        return result;
    }
}
