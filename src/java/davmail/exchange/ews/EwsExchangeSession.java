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
package davmail.exchange.ews;

import davmail.Settings;
import davmail.exception.DavMailAuthenticationException;
import davmail.exception.DavMailException;
import davmail.exception.HttpNotFoundException;
import davmail.exchange.*;
import davmail.exchange.condition.Condition;
import davmail.exchange.entity.*;
import davmail.http.DavGatewayHttpClientFacade;
import davmail.util.StringUtil;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.httpclient.*;
import org.apache.commons.httpclient.methods.ByteArrayRequestEntity;
import org.apache.commons.httpclient.methods.GetMethod;
import org.apache.commons.httpclient.methods.HeadMethod;
import org.apache.commons.httpclient.methods.PostMethod;
import org.apache.commons.httpclient.params.HttpClientParams;

import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * EWS Exchange adapter.
 * Compatible with Exchange 2007 and hopefully 2010.
 */
public class EwsExchangeSession extends ExchangeSession {

    protected static final int PAGE_SIZE = 100;

    protected static final String ARCHIVE_ROOT = "/archive/";

    /**
     * Message types.
     *
     * @see <a href="http://msdn.microsoft.com/en-us/library/aa565652%28v=EXCHG.140%29.aspx">http://msdn.microsoft.com/en-us/library/aa565652%28v=EXCHG.140%29.aspx</a>
     */
    protected static final Set<String> MESSAGE_TYPES = new HashSet<String>();

    static {
        MESSAGE_TYPES.add("Message");
        MESSAGE_TYPES.add("CalendarItem");

        MESSAGE_TYPES.add("MeetingMessage");
        MESSAGE_TYPES.add("MeetingRequest");
        MESSAGE_TYPES.add("MeetingResponse");
        MESSAGE_TYPES.add("MeetingCancellation");

        // exclude types from IMAP
        //MESSAGE_TYPES.add("Item");
        //MESSAGE_TYPES.add("PostItem");
        //MESSAGE_TYPES.add("Contact");
        //MESSAGE_TYPES.add("DistributionList");
        //MESSAGE_TYPES.add("Task");

        //ReplyToItem
        //ForwardItem
        //ReplyAllToItem
        //AcceptItem
        //TentativelyAcceptItem
        //DeclineItem
        //CancelCalendarItem
        //RemoveItem
        //PostReplyItem
        //SuppressReadReceipt
        //AcceptSharingInvitation
    }

    static final Map<String, String> vTodoToTaskStatusMap = new HashMap<String, String>();
    static final Map<String, String> taskTovTodoStatusMap = new HashMap<String, String>();

    static {
        //taskTovTodoStatusMap.put("NotStarted", null);
        taskTovTodoStatusMap.put("InProgress", "IN-PROCESS");
        taskTovTodoStatusMap.put("Completed", "COMPLETED");
        taskTovTodoStatusMap.put("WaitingOnOthers", "NEEDS-ACTION");
        taskTovTodoStatusMap.put("Deferred", "CANCELLED");

        //vTodoToTaskStatusMap.put(null, "NotStarted");
        vTodoToTaskStatusMap.put("IN-PROCESS", "InProgress");
        vTodoToTaskStatusMap.put("COMPLETED", "Completed");
        vTodoToTaskStatusMap.put("NEEDS-ACTION", "WaitingOnOthers");
        vTodoToTaskStatusMap.put("CANCELLED", "Deferred");
    }

    protected Map<String, String> folderIdMap;

    protected static class Folder extends davmail.exchange.entity.Folder {
        public FolderId folderId;

        public Folder(ExchangeSession exchangeSession) {
            super(exchangeSession);
        }
    }

    protected static class FolderPath {
        protected final String parentPath;
        protected final String folderName;

        protected FolderPath(String folderPath) {
            int slashIndex = folderPath.lastIndexOf('/');
            if (slashIndex < 0) {
                parentPath = "";
                folderName = folderPath;
            } else {
                parentPath = folderPath.substring(0, slashIndex);
                folderName = folderPath.substring(slashIndex + 1);
            }
        }
    }

    /**
     * @inheritDoc
     */
    public EwsExchangeSession(String url, String userName, String password) throws IOException {
        super(url, userName, password);
    }

    @Override
    protected HttpMethod formLogin(HttpClient httpClient, HttpMethod initmethod, String userName, String password) throws IOException {
        LOGGER.debug("Form based authentication detected");

        HttpMethod logonMethod = buildLogonMethod(httpClient, initmethod);
        if (logonMethod == null) {
            LOGGER.debug("Authentication form not found at " + initmethod.getURI() + ", will try direct EWS access");
        } else {
            logonMethod = postLogonMethod(httpClient, logonMethod, userName, password);
        }

        return logonMethod;
    }


    /**
     * Check endpoint url.
     *
     * @param endPointUrl endpoint url
     * @throws IOException on error
     */
    protected void checkEndPointUrl(String endPointUrl) throws IOException {
        HttpMethod checkMethod = new HeadMethod(endPointUrl);
        checkMethod.setFollowRedirects(true);
        try {
            int status = DavGatewayHttpClientFacade.executeNoRedirect(httpClient, checkMethod);
            if (status == HttpStatus.SC_UNAUTHORIZED) {
                throw new DavMailAuthenticationException("EXCEPTION_AUTHENTICATION_FAILED");
            } else if (status != HttpStatus.SC_OK) {
                throw new IOException("Ews endpoint not available at " + checkMethod.getURI().toString()+" status "+status);
            }
        } finally {
            checkMethod.releaseConnection();
        }
    }

    @Override
    protected void buildSessionInfo(HttpMethod method) throws DavMailException {
        // no need to check logon method body
        if (method != null) {
            method.releaseConnection();
        }
        boolean directEws = method == null || "/ews/services.wsdl".equalsIgnoreCase(method.getPath());

        // options page is not available in direct EWS mode
        if (!directEws) {
            // retrieve email and alias from options page
            getEmailAndAliasFromOptions();
        }

        if (email == null || alias == null) {
            // OWA authentication failed, get email address from login
            if (userName.indexOf('@') >= 0) {
                // userName is email address
                email = userName;
                alias = userName.substring(0, userName.indexOf('@'));
            } else {
                // userName or domain\\username, rebuild email address
                alias = getAliasFromLogin();
                email = getAliasFromLogin() + getEmailSuffixFromHostname();
            }
        }

        currentMailboxPath = "/users/" + email.toLowerCase();

        // check EWS access
        try {
            checkEndPointUrl("/ews/exchange.asmx");
            // workaround for Exchange bug: send fake request
            internalGetFolder("");
        } catch (IOException e) {
            // first failover: retry with NTLM
            DavGatewayHttpClientFacade.addNTLM(httpClient);
            try {
                checkEndPointUrl("/ews/exchange.asmx");
                // workaround for Exchange bug: send fake request
                internalGetFolder("");
            } catch (IOException e2) {
                LOGGER.debug(e2.getMessage());
                try {
                    // failover, try to retrieve EWS url from autodiscover
                    checkEndPointUrl(getEwsUrlFromAutoDiscover());
                    // workaround for Exchange bug: send fake request
                    internalGetFolder("");
                } catch (IOException e3) {
                    // autodiscover failed and initial exception was authentication failure => throw original exception
                    if (e instanceof DavMailAuthenticationException) {
                        throw (DavMailAuthenticationException) e;
                    }
                    LOGGER.error(e2.getMessage());
                    throw new DavMailAuthenticationException("EXCEPTION_EWS_NOT_AVAILABLE");
                }
            }
        }

        // enable preemptive authentication on non NTLM endpoints
        if (!DavGatewayHttpClientFacade.hasNTLMorNegotiate(httpClient)) {
            httpClient.getParams().setParameter(HttpClientParams.PREEMPTIVE_AUTHENTICATION, true);
        }

        // direct EWS: get primary smtp email address with ResolveNames
        if (directEws) {
            try {
                ResolveNamesMethod resolveNamesMethod = new ResolveNamesMethod(alias);
                executeMethod(resolveNamesMethod);
                List<EWSMethod.Item> responses = resolveNamesMethod.getResponseItems();
                for (EWSMethod.Item response : responses) {
                    if (alias.equalsIgnoreCase(response.get("Name"))) {
                        email = response.get("EmailAddress");
                        currentMailboxPath = "/users/" + email.toLowerCase();
                    }
                }
            } catch (IOException e) {
                LOGGER.warn("Unable to get primary email address with ResolveNames", e);
            }
        }

        try {
            folderIdMap = new HashMap<String, String>();
            // load actual well known folder ids
            folderIdMap.put(internalGetFolder(INBOX).folderId.value, INBOX);
            folderIdMap.put(internalGetFolder(CALENDAR).folderId.value, CALENDAR);
            folderIdMap.put(internalGetFolder(CONTACTS).folderId.value, CONTACTS);
            folderIdMap.put(internalGetFolder(SENT).folderId.value, SENT);
            folderIdMap.put(internalGetFolder(DRAFTS).folderId.value, DRAFTS);
            folderIdMap.put(internalGetFolder(TRASH).folderId.value, TRASH);
            folderIdMap.put(internalGetFolder(JUNK).folderId.value, JUNK);
            folderIdMap.put(internalGetFolder(UNSENT).folderId.value, UNSENT);
        } catch (IOException e) {
            LOGGER.error(e.getMessage(), e);
            throw new DavMailAuthenticationException("EXCEPTION_EWS_NOT_AVAILABLE");
        }
        LOGGER.debug("Current user email is " + email + ", alias is " + alias + " on " + serverVersion);
    }

    protected static class AutoDiscoverMethod extends PostMethod {
        AutoDiscoverMethod(String autodiscoverHost, String userEmail) {
            super("https://" + autodiscoverHost + "/autodiscover/autodiscover.xml");
            setAutoDiscoverRequestEntity(userEmail);
        }

        AutoDiscoverMethod(String userEmail) {
            super("/autodiscover/autodiscover.xml");
            setAutoDiscoverRequestEntity(userEmail);
        }

        void setAutoDiscoverRequestEntity(String userEmail) {
            String body = "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006\">" +
                    "<Request>" +
                    "<EMailAddress>" + userEmail + "</EMailAddress>" +
                    "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>" +
                    "</Request>" +
                    "</Autodiscover>";
            setRequestEntity(new ByteArrayRequestEntity(body.getBytes(), "text/xml"));
        }

        String ewsUrl;

        @Override
        protected void processResponseBody(HttpState httpState, HttpConnection httpConnection) {
            Header contentTypeHeader = getResponseHeader("Content-Type");
            if (contentTypeHeader != null &&
                    ("text/xml; charset=utf-8".equals(contentTypeHeader.getValue())
                            || "text/html; charset=utf-8".equals(contentTypeHeader.getValue())
                    )) {
                BufferedReader autodiscoverReader = null;
                try {
                    autodiscoverReader = new BufferedReader(new InputStreamReader(getResponseBodyAsStream()));
                    String line;
                    // find ews url
                    while ((line = autodiscoverReader.readLine()) != null
                            && (line.indexOf("<EwsUrl>") == -1)
                            && (line.indexOf("</EwsUrl>") == -1)) {
                    }
                    if (line != null) {
                        ewsUrl = line.substring(line.indexOf("<EwsUrl>") + 8, line.indexOf("</EwsUrl>"));
                    }
                } catch (IOException e) {
                    LOGGER.debug(e);
                } finally {
                    if (autodiscoverReader != null) {
                        try {
                            autodiscoverReader.close();
                        } catch (IOException e) {
                            LOGGER.debug(e);
                        }
                    }
                }
            }
        }
    }

    protected String getEwsUrlFromAutoDiscover() throws DavMailAuthenticationException {
        String ewsUrl;
        try {
            ewsUrl = getEwsUrlFromAutoDiscover(null);
        } catch (IOException e) {
            try {
                ewsUrl = getEwsUrlFromAutoDiscover("autodiscover." + email.substring(email.indexOf('@') + 1));
            } catch (IOException e2) {
                LOGGER.error(e2.getMessage());
                throw new DavMailAuthenticationException("EXCEPTION_EWS_NOT_AVAILABLE");
            }
        }
        return ewsUrl;
    }

    protected String getEwsUrlFromAutoDiscover(String autodiscoverHostname) throws IOException {
        String ewsUrl;
        AutoDiscoverMethod autoDiscoverMethod;
        if (autodiscoverHostname != null) {
            autoDiscoverMethod = new AutoDiscoverMethod(autodiscoverHostname, email);
        } else {
            autoDiscoverMethod = new AutoDiscoverMethod(email);
        }
        try {
            int status = DavGatewayHttpClientFacade.executeNoRedirect(httpClient, autoDiscoverMethod);
            if (status != HttpStatus.SC_OK) {
                throw DavGatewayHttpClientFacade.buildHttpException(autoDiscoverMethod);
            }
            ewsUrl = autoDiscoverMethod.ewsUrl;

            // update host name
            DavGatewayHttpClientFacade.setClientHost(httpClient, ewsUrl);

            if (ewsUrl == null) {
                throw new IOException("Ews url not found");
            }
        } finally {
            autoDiscoverMethod.releaseConnection();
        }
        return ewsUrl;
    }

    /**
     * Message create/update properties
     *
     * @param properties flag values map
     * @return field values
     */
    protected List<FieldUpdate> buildProperties(Map<String, String> properties) {
        ArrayList<FieldUpdate> list = new ArrayList<FieldUpdate>();
        for (Map.Entry<String, String> entry : properties.entrySet()) {
            if ("read".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("read", Boolean.toString("1".equals(entry.getValue()))));
            } else if ("junk".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("junk", entry.getValue()));
            } else if ("flagged".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("flagStatus", entry.getValue()));
            } else if ("answered".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("lastVerbExecuted", entry.getValue()));
                if ("102".equals(entry.getValue())) {
                    list.add(Field.createFieldUpdate("iconIndex", "261"));
                }
            } else if ("forwarded".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("lastVerbExecuted", entry.getValue()));
                if ("104".equals(entry.getValue())) {
                    list.add(Field.createFieldUpdate("iconIndex", "262"));
                }
            } else if ("draft".equals(entry.getKey())) {
                // note: draft is readonly after create
                list.add(Field.createFieldUpdate("messageFlags", entry.getValue()));
            } else if ("deleted".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("deleted", entry.getValue()));
            } else if ("datereceived".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("datereceived", entry.getValue()));
            } else if ("keywords".equals(entry.getKey())) {
                list.add(Field.createFieldUpdate("keywords", entry.getValue()));
            }
        }
        return list;
    }

    @Override
    public void createMessage(String folderPath, String messageName, HashMap<String, String> properties, MimeMessage mimeMessage) throws IOException {
        EWSMethod.Item item = new EWSMethod.Item();
        item.type = "Message";
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            mimeMessage.writeTo(baos);
        } catch (MessagingException e) {
            throw new IOException(e.getMessage());
        }
        baos.close();
        item.mimeContent = Base64.encodeBase64(baos.toByteArray());

        List<FieldUpdate> fieldUpdates = buildProperties(properties);
        if (!properties.containsKey("draft")) {
            // need to force draft flag to false
            if (properties.containsKey("read")) {
                fieldUpdates.add(Field.createFieldUpdate("messageFlags", "1"));
            } else {
                fieldUpdates.add(Field.createFieldUpdate("messageFlags", "0"));
            }
        }
        fieldUpdates.add(Field.createFieldUpdate("urlcompname", messageName));
        item.setFieldUpdates(fieldUpdates);
        CreateItemMethod createItemMethod = new CreateItemMethod(MessageDisposition.SaveOnly, getFolderId(folderPath), item);
        executeMethod(createItemMethod);
    }

    @Override
    public void updateMessage(Message message, Map<String, String> properties) throws IOException {
        if (properties.containsKey("read") && "urn:content-classes:appointment".equals(message.contentClass)) {
            properties.remove("read");
        }
        if (!properties.isEmpty()) {
            UpdateItemMethod updateItemMethod = new UpdateItemMethod(MessageDisposition.SaveOnly,
                    ConflictResolution.AlwaysOverwrite,
                    SendMeetingInvitationsOrCancellations.SendToNone,
                    ((EwsMessage) message).itemId, buildProperties(properties));
            executeMethod(updateItemMethod);
        }
    }

    @Override
    public void deleteMessage(Message message) throws IOException {
        LOGGER.debug("Delete " + message.permanentUrl);
        DeleteItemMethod deleteItemMethod = new DeleteItemMethod(((EwsMessage) message).itemId, DeleteType.HardDelete, SendMeetingCancellations.SendToNone);
        executeMethod(deleteItemMethod);
    }


    protected void sendMessage(String itemClass, byte[] messageBody) throws IOException {
        EWSMethod.Item item = new EWSMethod.Item();
        item.type = "Message";
        item.mimeContent = Base64.encodeBase64(messageBody);
        if (itemClass != null) {
            item.put("ItemClass", itemClass);
        }

        MessageDisposition messageDisposition;
        if (Settings.getBooleanProperty("davmail.smtpSaveInSent", true)) {
            messageDisposition = MessageDisposition.SendAndSaveCopy;
        } else {
            messageDisposition = MessageDisposition.SendOnly;
        }

        CreateItemMethod createItemMethod = new CreateItemMethod(messageDisposition, getFolderId(SENT), item);
        executeMethod(createItemMethod);
    }

    @Override
    public void sendMessage(MimeMessage mimeMessage) throws IOException, MessagingException {
        String itemClass = null;
        if (mimeMessage.getContentType().startsWith("multipart/report")) {
            itemClass = "REPORT.IPM.Note.IPNRN";
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            mimeMessage.writeTo(baos);
        } catch (MessagingException e) {
            throw new IOException(e.getMessage());
        }
        sendMessage(itemClass, baos.toByteArray());
    }

    /**
     * @inheritDoc
     */
    @Override
    public byte[] getContent(Message message) throws IOException {
        return getContent(((EwsMessage) message).itemId);
    }

    /**
     * Get item content.
     *
     * @param itemId EWS item id
     * @return item content as byte array
     * @throws IOException on error
     */
    protected byte[] getContent(ItemId itemId) throws IOException {
        GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, true);
        byte[] mimeContent = null;
        try {
            executeMethod(getItemMethod);
            mimeContent = getItemMethod.getMimeContent();
        } catch (EWSException e) {
            LOGGER.warn("GetItem with MimeContent failed: " + e.getMessage());
        }
        if (getItemMethod.getStatusCode() == HttpStatus.SC_NOT_FOUND) {
            throw new HttpNotFoundException("Item " + itemId + " not found");
        }
        if (mimeContent == null) {
            LOGGER.warn("MimeContent not available, trying to rebuild from properties");
            try {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, false);
                getItemMethod.addAdditionalProperty(Field.get("contentclass"));
                getItemMethod.addAdditionalProperty(Field.get("message-id"));
                getItemMethod.addAdditionalProperty(Field.get("from"));
                getItemMethod.addAdditionalProperty(Field.get("to"));
                getItemMethod.addAdditionalProperty(Field.get("cc"));
                getItemMethod.addAdditionalProperty(Field.get("subject"));
                getItemMethod.addAdditionalProperty(Field.get("date"));
                getItemMethod.addAdditionalProperty(Field.get("body"));
                executeMethod(getItemMethod);
                EWSMethod.Item item = getItemMethod.getResponseItem();

                if (item == null) {
                    throw new HttpNotFoundException("Item " + itemId + " not found");
                }

                MimeMessage mimeMessage = new MimeMessage((Session) null);
                mimeMessage.addHeader("Content-class", item.get(Field.get("contentclass").getResponseName()));
                mimeMessage.setSentDate(parseDateFromExchange(item.get(Field.get("date").getResponseName())));
                mimeMessage.addHeader("From", item.get(Field.get("from").getResponseName()));
                mimeMessage.addHeader("To", item.get(Field.get("to").getResponseName()));
                mimeMessage.addHeader("Cc", item.get(Field.get("cc").getResponseName()));
                mimeMessage.setSubject(item.get(Field.get("subject").getResponseName()));
                String propertyValue = item.get(Field.get("body").getResponseName());
                if (propertyValue == null) {
                    propertyValue = "";
                }
                mimeMessage.setContent(propertyValue, "text/html; charset=UTF-8");

                mimeMessage.writeTo(baos);
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug("Rebuilt message content: " + new String(baos.toByteArray()));
                }
                mimeContent = baos.toByteArray();

            } catch (IOException e2) {
                LOGGER.warn(e2);
            } catch (MessagingException e2) {
                LOGGER.warn(e2);
            }
            if (mimeContent == null) {
                throw new IOException("GetItem returned null MimeContent");
            }
        }
        return mimeContent;
    }

    protected EwsMessage buildMessage(EWSMethod.Item response) throws DavMailException {
        EwsMessage message = new EwsMessage(this);

        // get item id
        message.itemId = new ItemId(response);

        message.permanentUrl = response.get(Field.get("permanenturl").getResponseName());

        message.size = response.getInt(Field.get("messageSize").getResponseName());
        message.uid = response.get(Field.get("uid").getResponseName());
        message.contentClass = response.get(Field.get("contentclass").getResponseName());
        message.imapUid = response.getLong(Field.get("imapUid").getResponseName());
        message.read = response.getBoolean(Field.get("read").getResponseName());
        message.junk = response.getBoolean(Field.get("junk").getResponseName());
        message.flagged = "2".equals(response.get(Field.get("flagStatus").getResponseName()));
        message.draft = (response.getInt(Field.get("messageFlags").getResponseName()) & 8) != 0;
        String lastVerbExecuted = response.get(Field.get("lastVerbExecuted").getResponseName());
        message.answered = "102".equals(lastVerbExecuted) || "103".equals(lastVerbExecuted);
        message.forwarded = "104".equals(lastVerbExecuted);
        message.date = convertDateFromExchange(response.get(Field.get("date").getResponseName()));
        message.deleted = "1".equals(response.get(Field.get("deleted").getResponseName()));

        String lastmodified = convertDateFromExchange(response.get(Field.get("lastmodified").getResponseName()));
        message.recent = !message.read && lastmodified != null && lastmodified.equals(message.date);

        message.keywords = response.get(Field.get("keywords").getResponseName());

        if (LOGGER.isDebugEnabled()) {
            StringBuilder buffer = new StringBuilder();
            buffer.append("Message");
            if (message.imapUid != 0) {
                buffer.append(" IMAP uid: ").append(message.imapUid);
            }
            if (message.uid != null) {
                buffer.append(" uid: ").append(message.uid);
            }
            buffer.append(" ItemId: ").append(message.itemId.id);
            buffer.append(" ChangeKey: ").append(message.itemId.changeKey);
            LOGGER.debug(buffer.toString());
        }
        return message;
    }

    @Override
    public MessageList searchMessages(String folderPath, Set<String> attributes, Condition condition) throws IOException {
        MessageList messages = new MessageList();
        int maxCount = Settings.getIntProperty("davmail.folderSizeLimit", 0);
        List<EWSMethod.Item> responses = searchItems(folderPath, attributes, condition, FolderQueryTraversal.SHALLOW, maxCount);

        for (EWSMethod.Item response : responses) {
            if (MESSAGE_TYPES.contains(response.type)) {
                EwsMessage message = buildMessage(response);
                message.messageList = messages;
                messages.add(message);
            }
        }
        Collections.sort(messages);
        return messages;
    }

    protected List<EWSMethod.Item> searchItems(String folderPath, Set<String> attributes, Condition condition, FolderQueryTraversal folderQueryTraversal, int maxCount) throws IOException {
        int offset = 0;
        List<EWSMethod.Item> results = new ArrayList<EWSMethod.Item>();
        FindItemMethod findItemMethod;
        do {
            int fetchCount = PAGE_SIZE;
            if (maxCount > 0) {
                fetchCount = Math.min(PAGE_SIZE, maxCount - offset);
            }
            findItemMethod = new FindItemMethod(folderQueryTraversal, BaseShape.ID_ONLY, getFolderId(folderPath), offset, fetchCount);
            for (String attribute : attributes) {
                findItemMethod.addAdditionalProperty(Field.get(attribute));
            }
            if (condition != null && !condition.isEmpty()) {
                findItemMethod.setSearchExpression((SearchExpression) condition);
            }
            findItemMethod.setFieldOrder(new FieldOrder(Field.get("imapUid"), FieldOrder.Order.Descending));
            executeMethod(findItemMethod);
            results.addAll(findItemMethod.getResponseItems());
            offset = results.size();
            if (Thread.interrupted()) {
                LOGGER.debug("Search items failed: Interrupted by client");
                throw new IOException("Search items failed: Interrupted by client");
            }
        } while (!(findItemMethod.includesLastItemInRange || (maxCount > 0 && offset == maxCount)));
        return results;
    }

    protected static class MultiCondition extends davmail.exchange.condition.MultiCondition implements SearchExpression {
        protected MultiCondition(Operator operator, Condition... condition) {
            super(operator, condition);
        }

        public void appendTo(StringBuilder buffer) {
            int actualConditionCount = 0;
            for (Condition condition : conditions) {
                if (!condition.isEmpty()) {
                    actualConditionCount++;
                }
            }
            if (actualConditionCount > 0) {
                if (actualConditionCount > 1) {
                    buffer.append("<t:").append(operator.toString()).append('>');
                }

                for (Condition condition : conditions) {
                    condition.appendTo(buffer);
                }

                if (actualConditionCount > 1) {
                    buffer.append("</t:").append(operator.toString()).append('>');
                }
            }
        }
    }

    protected static class NotCondition extends davmail.exchange.condition.NotCondition implements SearchExpression {
        protected NotCondition(Condition condition) {
            super(condition);
        }

        public void appendTo(StringBuilder buffer) {
            buffer.append("<t:Not>");
            condition.appendTo(buffer);
            buffer.append("</t:Not>");
        }
    }


    protected static class AttributeCondition extends davmail.exchange.condition.AttributeCondition implements SearchExpression {
        protected ContainmentMode containmentMode;
        protected ContainmentComparison containmentComparison;

        protected AttributeCondition(String attributeName, Operator operator, String value) {
            super(attributeName, operator, value);
        }

        protected AttributeCondition(String attributeName, Operator operator, String value,
                                     ContainmentMode containmentMode, ContainmentComparison containmentComparison) {
            super(attributeName, operator, value);
            this.containmentMode = containmentMode;
            this.containmentComparison = containmentComparison;
        }

        protected FieldURI getFieldURI() {
            return Field.get(attributeName);
        }

        protected Operator getOperator() {
            return operator;
        }

        public void appendTo(StringBuilder buffer) {
            buffer.append("<t:").append(operator.toString());
            if (containmentMode != null) {
                containmentMode.appendTo(buffer);
            }
            if (containmentComparison != null) {
                containmentComparison.appendTo(buffer);
            }
            buffer.append('>');
            FieldURI fieldURI = getFieldURI();
            fieldURI.appendTo(buffer);

            if (operator != Operator.Contains) {
                buffer.append("<t:FieldURIOrConstant>");
            }
            buffer.append("<t:Constant Value=\"");
            // encode urlcompname
            if (fieldURI instanceof ExtendedFieldURI && "0x10f3".equals(((ExtendedFieldURI) fieldURI).propertyTag)) {
                buffer.append(StringUtil.xmlEncodeAttribute(StringUtil.encodeUrlcompname(value)));
            } else if (fieldURI instanceof ExtendedFieldURI
                    && ((ExtendedFieldURI) fieldURI).propertyType == ExtendedFieldURI.PropertyType.Integer) {
                // check value
                try {
                    Integer.parseInt(value);
                    buffer.append(value);
                } catch (NumberFormatException e) {
                    // invalid value, replace with 0
                    buffer.append('0');
                }
            } else {
                buffer.append(StringUtil.xmlEncodeAttribute(value));
            }
            buffer.append("\"/>");
            if (operator != Operator.Contains) {
                buffer.append("</t:FieldURIOrConstant>");
            }

            buffer.append("</t:").append(operator.toString()).append('>');
        }

        public boolean isMatch(Contact contact) {
            String lowerCaseValue = value.toLowerCase();

            String actualValue = contact.get(attributeName);
            if (actualValue == null) {
                return false;
            }
            actualValue = actualValue.toLowerCase();
            if (operator == Operator.IsEqualTo) {
                return lowerCaseValue.equals(actualValue);
            } else {
                return operator == Operator.Contains && ((containmentMode.equals(ContainmentMode.Substring) && actualValue.contains(lowerCaseValue)) ||
                        (containmentMode.equals(ContainmentMode.Prefixed) && actualValue.startsWith(lowerCaseValue)));
            }
        }

    }

    @Override
    public MultiCondition and(Condition... condition) {
        return new MultiCondition(Operator.And, condition);
    }

    @Override
    public MultiCondition or(Condition... condition) {
        return new MultiCondition(Operator.Or, condition);
    }

    @Override
    public Condition not(Condition condition) {
        return new NotCondition(condition);
    }

    @Override
    public Condition isEqualTo(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.IsEqualTo, value);
    }

    @Override
    public Condition isEqualTo(String attributeName, int value) {
        return new AttributeCondition(attributeName, Operator.IsEqualTo, String.valueOf(value));
    }

    @Override
    public Condition headerIsEqualTo(String headerName, String value) {
        if (serverVersion.startsWith("Exchange2010")) {
            if ("from".equals(headerName)
                    || "to".equals(headerName)
                    || "cc".equals(headerName)) {
                return new AttributeCondition("msg"+headerName, Operator.Contains, value, ContainmentMode.Substring, ContainmentComparison.IgnoreCase);
            } else if ("message-id".equals(headerName)
                    || "bcc".equals(headerName)) {
                return new AttributeCondition(headerName, Operator.Contains, value, ContainmentMode.Substring, ContainmentComparison.IgnoreCase);
            } else {
                // Exchange 2010 does not support header search, use PR_TRANSPORT_MESSAGE_HEADERS instead
                return new AttributeCondition("messageheaders", Operator.Contains, headerName + ": " + value, ContainmentMode.Substring, ContainmentComparison.IgnoreCase);
            }
        } else {
            return new HeaderCondition(headerName, value);
        }
    }

    @Override
    public Condition gte(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.IsGreaterThanOrEqualTo, value);
    }

    @Override
    public Condition lte(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.IsLessThanOrEqualTo, value);
    }

    @Override
    public Condition lt(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.IsLessThan, value);
    }

    @Override
    public Condition gt(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.IsGreaterThan, value);
    }

    @Override
    public Condition contains(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.Contains, value, ContainmentMode.Substring, ContainmentComparison.IgnoreCase);
    }

    @Override
    public Condition startsWith(String attributeName, String value) {
        return new AttributeCondition(attributeName, Operator.Contains, value, ContainmentMode.Prefixed, ContainmentComparison.IgnoreCase);
    }

    @Override
    public Condition isNull(String attributeName) {
        return new IsNullCondition(attributeName);
    }

    @Override
    public Condition isTrue(String attributeName) {
        return new AttributeCondition(attributeName, Operator.IsEqualTo, "true");
    }

    @Override
    public Condition isFalse(String attributeName) {
        return new AttributeCondition(attributeName, Operator.IsEqualTo, "false");
    }

    protected static final HashSet<FieldURI> FOLDER_PROPERTIES = new HashSet<FieldURI>();

    static {
        FOLDER_PROPERTIES.add(Field.get("urlcompname"));
        FOLDER_PROPERTIES.add(Field.get("folderDisplayName"));
        FOLDER_PROPERTIES.add(Field.get("lastmodified"));
        FOLDER_PROPERTIES.add(Field.get("folderclass"));
        FOLDER_PROPERTIES.add(Field.get("ctag"));
        FOLDER_PROPERTIES.add(Field.get("count"));
        FOLDER_PROPERTIES.add(Field.get("unread"));
        FOLDER_PROPERTIES.add(Field.get("hassubs"));
        FOLDER_PROPERTIES.add(Field.get("uidNext"));
        FOLDER_PROPERTIES.add(Field.get("highestUid"));
    }

    protected Folder buildFolder(EWSMethod.Item item) {
        Folder folder = new Folder(this);
        folder.folderId = new FolderId(item);
        folder.displayName = item.get(Field.get("folderDisplayName").getResponseName());
        folder.folderClass = item.get(Field.get("folderclass").getResponseName());
        folder.etag = item.get(Field.get("lastmodified").getResponseName());
        folder.ctag = item.get(Field.get("ctag").getResponseName());
        folder.count = item.getInt(Field.get("count").getResponseName());
        folder.unreadCount = item.getInt(Field.get("unread").getResponseName());
        // fake recent value
        folder.recent = folder.unreadCount;
        folder.hasChildren = item.getBoolean(Field.get("hassubs").getResponseName());
        // noInferiors not implemented
        folder.uidNext = item.getInt(Field.get("uidNext").getResponseName());
        return folder;
    }

    /**
     * @inheritDoc
     */
    @Override
    public List<davmail.exchange.entity.Folder> getSubFolders(String folderPath, Condition condition, boolean recursive) throws IOException {
        String baseFolderPath = folderPath;
        if (baseFolderPath.startsWith("/users/")) {
            int index = baseFolderPath.indexOf('/', "/users/".length());
            if (index >= 0) {
                baseFolderPath = baseFolderPath.substring(index + 1);
            }
        }
        List<davmail.exchange.entity.Folder> folders = new ArrayList<davmail.exchange.entity.Folder>();
        appendSubFolders(folders, baseFolderPath, getFolderId(folderPath), condition, recursive);
        return folders;
    }

    protected void appendSubFolders(List<davmail.exchange.entity.Folder> folders,
                                    String parentFolderPath, FolderId parentFolderId,
                                    Condition condition, boolean recursive) throws IOException {
        FindFolderMethod findFolderMethod = new FindFolderMethod(FolderQueryTraversal.SHALLOW,
                BaseShape.ID_ONLY, parentFolderId, FOLDER_PROPERTIES, (SearchExpression) condition);
        executeMethod(findFolderMethod);
        for (EWSMethod.Item item : findFolderMethod.getResponseItems()) {
            Folder folder = buildFolder(item);
            if (parentFolderPath.length() > 0) {
                if (parentFolderPath.endsWith("/")) {
                    folder.folderPath = parentFolderPath + item.get(Field.get("folderDisplayName").getResponseName());
                } else {
                    folder.folderPath = parentFolderPath + '/' + item.get(Field.get("folderDisplayName").getResponseName());
                }
            } else if (folderIdMap.get(folder.folderId.value) != null) {
                folder.folderPath = folderIdMap.get(folder.folderId.value);
            } else {
                folder.folderPath = item.get(Field.get("folderDisplayName").getResponseName());
            }
            folders.add(folder);
            if (recursive && folder.hasChildren) {
                appendSubFolders(folders, folder.folderPath, folder.folderId, condition, recursive);
            }
        }
    }

    /**
     * Get folder by path.
     *
     * @param folderPath folder path
     * @return folder object
     * @throws IOException on error
     */
    @Override
    protected EwsExchangeSession.Folder internalGetFolder(String folderPath) throws IOException {
        FolderId folderId = getFolderId(folderPath);
        GetFolderMethod getFolderMethod = new GetFolderMethod(BaseShape.ID_ONLY, folderId, FOLDER_PROPERTIES);
        executeMethod(getFolderMethod);
        EWSMethod.Item item = getFolderMethod.getResponseItem();
        Folder folder;
        if (item != null) {
            folder = buildFolder(item);
            folder.folderPath = folderPath;
        } else {
            throw new HttpNotFoundException("Folder " + folderPath + " not found");
        }
        return folder;
    }

    /**
     * @inheritDoc
     */
    @Override
    public int createFolder(String folderPath, String folderClass, Map<String, String> properties) throws IOException {
        FolderPath path = new FolderPath(folderPath);
        EWSMethod.Item folder = new EWSMethod.Item();
        folder.type = "Folder";
        folder.put("FolderClass", folderClass);
        folder.put("DisplayName", path.folderName);
        // TODO: handle properties
        CreateFolderMethod createFolderMethod = new CreateFolderMethod(getFolderId(path.parentPath), folder);
        executeMethod(createFolderMethod);
        return HttpStatus.SC_CREATED;
    }

    /**
     * @inheritDoc
     */
    @Override
    public int updateFolder(String folderPath, Map<String, String> properties) throws IOException {
        ArrayList<FieldUpdate> updates = new ArrayList<FieldUpdate>();
        for (Map.Entry<String, String> entry : properties.entrySet()) {
            updates.add(new FieldUpdate(Field.get(entry.getKey()), entry.getValue()));
        }
        UpdateFolderMethod updateFolderMethod = new UpdateFolderMethod(internalGetFolder(folderPath).folderId, updates);

        executeMethod(updateFolderMethod);
        return HttpStatus.SC_CREATED;
    }

    /**
     * @inheritDoc
     */
    @Override
    public void deleteFolder(String folderPath) throws IOException {
        FolderId folderId = getFolderIdIfExists(folderPath);
        if (folderId != null) {
            DeleteFolderMethod deleteFolderMethod = new DeleteFolderMethod(folderId);
            executeMethod(deleteFolderMethod);
        } else {
            LOGGER.debug("Folder " + folderPath + " not found");
        }
    }

    /**
     * @inheritDoc
     */
    @Override
    public void moveMessage(Message message, String targetFolder) throws IOException {
        MoveItemMethod moveItemMethod = new MoveItemMethod(((EwsMessage) message).itemId, getFolderId(targetFolder));
        executeMethod(moveItemMethod);
    }

    /**
     * @inheritDoc
     */
    @Override
    public void copyMessage(Message message, String targetFolder) throws IOException {
        CopyItemMethod copyItemMethod = new CopyItemMethod(((EwsMessage) message).itemId, getFolderId(targetFolder));
        executeMethod(copyItemMethod);
    }

    /**
     * @inheritDoc
     */
    @Override
    public void moveFolder(String folderPath, String targetFolderPath) throws IOException {
        FolderPath path = new FolderPath(folderPath);
        FolderPath targetPath = new FolderPath(targetFolderPath);
        FolderId folderId = getFolderId(folderPath);
        FolderId toFolderId = getFolderId(targetPath.parentPath);
        toFolderId.changeKey = null;
        // move folder
        if (!path.parentPath.equals(targetPath.parentPath)) {
            MoveFolderMethod moveFolderMethod = new MoveFolderMethod(folderId, toFolderId);
            executeMethod(moveFolderMethod);
        }
        // rename folder
        if (!path.folderName.equals(targetPath.folderName)) {
            ArrayList<FieldUpdate> updates = new ArrayList<FieldUpdate>();
            updates.add(new FieldUpdate(Field.get("folderDisplayName"), targetPath.folderName));
            UpdateFolderMethod updateFolderMethod = new UpdateFolderMethod(folderId, updates);
            executeMethod(updateFolderMethod);
        }
    }

    @Override
    public void moveItem(String sourcePath, String targetPath) throws IOException {
        FolderPath sourceFolderPath = new FolderPath(sourcePath);
        Item item = getItem(sourceFolderPath.parentPath, sourceFolderPath.folderName);
        FolderPath targetFolderPath = new FolderPath(targetPath);
        FolderId toFolderId = getFolderId(targetFolderPath.parentPath);
        MoveItemMethod moveItemMethod = new MoveItemMethod(((EwsEvent) item).itemId, toFolderId);
        executeMethod(moveItemMethod);
    }

    /**
     * @inheritDoc
     */
    @Override
    public void moveToTrash(Message message) throws IOException {
        MoveItemMethod moveItemMethod = new MoveItemMethod(((EwsMessage) message).itemId, getFolderId(TRASH));
        executeMethod(moveItemMethod);
    }

    @Override
    public List<Contact> searchContacts(String folderPath, Set<String> attributes, Condition condition, int maxCount) throws IOException {
        List<EWSMethod.Item> responses = searchItems(folderPath, attributes, condition,
                FolderQueryTraversal.SHALLOW, maxCount);

        List<Contact> contacts = new ArrayList<Contact>(responses.size());
        for (EWSMethod.Item response : responses) {
            contacts.add(new EwsContact(this, response));
        }
        return contacts;
    }

    @Override
    protected Condition getCalendarItemCondition(Condition dateCondition) {
        // tasks in calendar not supported over EWS => do not look for instancetype null
        return or(
                // Exchange 2010
                or(isTrue("isrecurring"),
                        and(isFalse("isrecurring"), dateCondition)),
                // Exchange 2007
                or(isEqualTo("instancetype", 1),
                        and(isEqualTo("instancetype", 0), dateCondition))
        );
    }

    @Override
    public List<Event> getEventMessages(String folderPath) throws IOException {
        return searchEvents(folderPath, ITEM_PROPERTIES,
                and(startsWith("outlookmessageclass", "IPM.Schedule.Meeting."),
                        or(isNull("processed"), isFalse("processed"))));
    }

    @Override
    public List<Event> searchEvents(String folderPath, Set<String> attributes, Condition condition) throws IOException {
        List<Event> events = new ArrayList<Event>();
        List<EWSMethod.Item> responses = searchItems(folderPath, attributes,
                condition,
                FolderQueryTraversal.SHALLOW, 0);
        for (EWSMethod.Item response : responses) {
            EwsEvent event = new EwsEvent(this, response);
            if ("Message".equals(event.type)) {
                // TODO: just exclude
                // need to check body
                try {
                    event.getEventContent();
                    events.add(event);
                } catch (HttpException e) {
                    LOGGER.warn("Ignore invalid event " + event.getHref());
                }
                // exclude exceptions
            } else if (event.isException) {
                LOGGER.debug("Exclude recurrence exception " + event.getHref());
            } else {
                events.add(event);
            }

        }

        return events;
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

        ITEM_PROPERTIES.add("calendaritemtype");
        ITEM_PROPERTIES.add("isrecurring");
    }

    protected static final HashSet<String> EVENT_REQUEST_PROPERTIES = new HashSet<String>();

    static {
        EVENT_REQUEST_PROPERTIES.add("permanenturl");
        EVENT_REQUEST_PROPERTIES.add("etag");
        EVENT_REQUEST_PROPERTIES.add("displayname");
        EVENT_REQUEST_PROPERTIES.add("subject");
        EVENT_REQUEST_PROPERTIES.add("urlcompname");
    }

    protected Set<String> getItemProperties() {
        return ITEM_PROPERTIES;
    }

    protected EWSMethod.Item getEwsItem(String folderPath, String itemName) throws IOException {
        EWSMethod.Item item = null;
        String urlcompname = convertItemNameToEML(itemName);
        // workaround for missing urlcompname in Exchange 2010
        if (isItemId(urlcompname)) {
            ItemId itemId = new ItemId(StringUtil.urlToBase64(urlcompname.substring(0, urlcompname.indexOf('.'))));
            GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, false);
            for (String attribute : EVENT_REQUEST_PROPERTIES) {
                getItemMethod.addAdditionalProperty(Field.get(attribute));
            }
            executeMethod(getItemMethod);
            item = getItemMethod.getResponseItem();
        }
        // find item by urlcompname
        if (item == null) {
            List<EWSMethod.Item> responses = searchItems(folderPath, EVENT_REQUEST_PROPERTIES, isEqualTo("urlcompname", urlcompname), FolderQueryTraversal.SHALLOW, 0);
            if (!responses.isEmpty()) {
                item = responses.get(0);
            }
        }
        return item;
    }


    @Override
    public Item getItem(String folderPath, String itemName) throws IOException {
        EWSMethod.Item item = getEwsItem(folderPath, itemName);
        if (item == null && isMainCalendar(folderPath)) {
            // look for item in task folder, replace extension first
            if (itemName.endsWith(".ics")) {
                itemName = itemName.substring(0, itemName.length() - 3) + "EML";
            }
            item = getEwsItem(TASKS, itemName);
        }

        if (item == null) {
            throw new HttpNotFoundException(itemName + " not found in " + folderPath);
        }

        String itemType = item.type;
        if ("Contact".equals(itemType)) {
            // retrieve Contact properties
            ItemId itemId = new ItemId(item);
            GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, false);
            for (String attribute : CONTACT_ATTRIBUTES) {
                getItemMethod.addAdditionalProperty(Field.get(attribute));
            }
            executeMethod(getItemMethod);
            item = getItemMethod.getResponseItem();
            if (item == null) {
                throw new HttpNotFoundException(itemName + " not found in " + folderPath);
            }
            return new EwsContact(this, item);
        } else if ("CalendarItem".equals(itemType)
                || "MeetingRequest".equals(itemType)
                || "Task".equals(itemType)
                // VTODOs appear as Messages
                || "Message".equals(itemType)) {
            return new EwsEvent(this, item);
        } else {
            throw new HttpNotFoundException(itemName + " not found in " + folderPath);
        }

    }

    @Override
    public ContactPhoto getContactPhoto(Contact contact) throws IOException {
        ContactPhoto contactPhoto = null;

        GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, ((EwsContact) contact).itemId, false);
        getItemMethod.addAdditionalProperty(Field.get("attachments"));
        executeMethod(getItemMethod);
        EWSMethod.Item item = getItemMethod.getResponseItem();
        if (item != null) {
            FileAttachment attachment = item.getAttachmentByName("ContactPicture.jpg");
            if (attachment == null) {
                throw new IOException("Missing contact picture");
            }
            // get attachment content
            GetAttachmentMethod getAttachmentMethod = new GetAttachmentMethod(attachment.attachmentId);
            executeMethod(getAttachmentMethod);

            contactPhoto = new ContactPhoto();
            contactPhoto.content = getAttachmentMethod.getResponseItem().get("Content");
            if (attachment.contentType == null) {
                contactPhoto.contentType = "image/jpeg";
            } else {
                contactPhoto.contentType = attachment.contentType;
            }
        }

        return contactPhoto;
    }

    @Override
    public void deleteItem(String folderPath, String itemName) throws IOException {
        EWSMethod.Item item = getEwsItem(folderPath, itemName);
        if (item == null && isMainCalendar(folderPath)) {
            // look for item in task folder
            item = getEwsItem(TASKS, itemName);
        }
        if (item != null) {
            DeleteItemMethod deleteItemMethod = new DeleteItemMethod(new ItemId(item), DeleteType.HardDelete, SendMeetingCancellations.SendToNone);
            executeMethod(deleteItemMethod);
        }
    }

    @Override
    public void processItem(String folderPath, String itemName) throws IOException {
        EWSMethod.Item item = getEwsItem(folderPath, itemName);
        if (item != null) {
            HashMap<String, String> localProperties = new HashMap<String, String>();
            localProperties.put("processed", "1");
            localProperties.put("read", "1");
            UpdateItemMethod updateItemMethod = new UpdateItemMethod(MessageDisposition.SaveOnly,
                    ConflictResolution.AlwaysOverwrite,
                    SendMeetingInvitationsOrCancellations.SendToNone,
                    new ItemId(item), buildProperties(localProperties));
            executeMethod(updateItemMethod);
        }
    }

    @Override
    public int sendEvent(String icsBody) throws IOException {
        String itemName = UUID.randomUUID().toString() + ".EML";
        byte[] mimeContent = new EwsEvent(this, DRAFTS, itemName, "urn:content-classes:calendarmessage", icsBody, null, null).createMimeContent();
        if (mimeContent == null) {
            // no recipients, cancel
            return HttpStatus.SC_NO_CONTENT;
        } else {
            sendMessage(null, mimeContent);
            return HttpStatus.SC_OK;
        }
    }

    @Override
    protected ItemResult internalCreateOrUpdateContact(String folderPath, String itemName, Map<String, String> properties, String etag, String noneMatch) throws IOException {
        return new EwsContact(this, folderPath, itemName, properties, StringUtil.removeQuotes(etag), noneMatch).createOrUpdate();
    }

    @Override
    protected ItemResult internalCreateOrUpdateEvent(String folderPath, String itemName, String contentClass, String icsBody, String etag, String noneMatch) throws IOException {
        return new EwsEvent(this, folderPath, itemName, contentClass, icsBody, StringUtil.removeQuotes(etag), noneMatch).createOrUpdate();
    }

    @Override
    public boolean isSharedFolder(String folderPath) {
        return folderPath.startsWith("/") && !folderPath.toLowerCase().startsWith(currentMailboxPath);
    }

    @Override
    public boolean isMainCalendar(String folderPath) {
        return "calendar".equalsIgnoreCase(folderPath) || (currentMailboxPath + "/calendar").equalsIgnoreCase(folderPath);
    }

    @Override
    protected String getFreeBusyData(String attendee, String start, String end, int interval) throws IOException {
        GetUserAvailabilityMethod getUserAvailabilityMethod = new GetUserAvailabilityMethod(attendee, start, end, interval);
        executeMethod(getUserAvailabilityMethod);
        return getUserAvailabilityMethod.getMergedFreeBusy();
    }

    @Override
    protected void loadVtimezone() {

        try {
            String timezoneId = null;
            if (!"Exchange2007_SP1".equals(serverVersion)) {
                GetUserConfigurationMethod getUserConfigurationMethod = new GetUserConfigurationMethod();
                executeMethod(getUserConfigurationMethod);
                EWSMethod.Item item = getUserConfigurationMethod.getResponseItem();
                if (item != null) {
                    timezoneId = item.get("timezone");
                }
            } else {
                timezoneId = getTimezoneidFromOptions();
            }
            // failover: use timezone id from settings file
            if (timezoneId == null) {
                timezoneId = Settings.getProperty("davmail.timezoneId");
            }
            // last failover: use GMT
            if (timezoneId == null) {
                LOGGER.warn("Unable to get user timezone, using GMT Standard Time. Set davmail.timezoneId setting to override this.");
                timezoneId = "GMT Standard Time";
            }

            createCalendarFolder("davmailtemp", null);
            EWSMethod.Item item = new EWSMethod.Item();
            item.type = "CalendarItem";
            if (!"Exchange2007_SP1".equals(serverVersion)) {
                SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss", Locale.ENGLISH);
                dateFormatter.setTimeZone(GMT_TIMEZONE);
                Calendar cal = Calendar.getInstance();
                item.put("Start", dateFormatter.format(cal.getTime()));
                cal.add(Calendar.DAY_OF_MONTH, 1);
                item.put("End", dateFormatter.format(cal.getTime()));
                item.put("StartTimeZone", timezoneId);
            } else {
                item.put("MeetingTimeZone", timezoneId);
            }
            CreateItemMethod createItemMethod = new CreateItemMethod(MessageDisposition.SaveOnly, SendMeetingInvitations.SendToNone, getFolderId("davmailtemp"), item);
            executeMethod(createItemMethod);
            item = createItemMethod.getResponseItem();
            VCalendar vCalendar = new VCalendar(getContent(new ItemId(item)), email, null);
            this.vTimezone = vCalendar.getVTimezone();
            // delete temporary folder
            deleteFolder("davmailtemp");
        } catch (IOException e) {
            LOGGER.warn("Unable to get VTIMEZONE info: " + e, e);
        }
    }

    protected String getTimezoneidFromOptions() {
        String result = null;
        // get user mail URL from html body
        BufferedReader optionsPageReader = null;
        GetMethod optionsMethod = new GetMethod("/owa/?ae=Options&t=Regional");
        try {
            DavGatewayHttpClientFacade.executeGetMethod(httpClient, optionsMethod, false);
            optionsPageReader = new BufferedReader(new InputStreamReader(optionsMethod.getResponseBodyAsStream()));
            String line;
            // find email
            //noinspection StatementWithEmptyBody
            while ((line = optionsPageReader.readLine()) != null
                    && (line.indexOf("tblTmZn") == -1)
                    && (line.indexOf("selTmZn") == -1)) {
            }
            if (line != null) {
                if (line.indexOf("tblTmZn") >= 0) {
                    int start = line.indexOf("oV=\"") + 4;
                    int end = line.indexOf('\"', start);
                    result = line.substring(start, end);
                } else {
                    int end = line.lastIndexOf("\" selected>");
                    int start = line.lastIndexOf('\"', end - 1);
                    result = line.substring(start + 1, end);
                }
            }
        } catch (IOException e) {
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

        return result;
    }


    protected FolderId getFolderId(String folderPath) throws IOException {
        FolderId folderId = getFolderIdIfExists(folderPath);
        if (folderId == null) {
            throw new HttpNotFoundException("Folder '" + folderPath + "' not found");
        }
        return folderId;
    }

    protected static final String USERS_ROOT = "/users/";

    protected FolderId getFolderIdIfExists(String folderPath) throws IOException {
        String lowerCaseFolderPath = folderPath.toLowerCase();
        if (lowerCaseFolderPath.equals(currentMailboxPath)) {
            return getSubFolderIdIfExists(null, "");
        } else if (lowerCaseFolderPath.startsWith(currentMailboxPath + '/')) {
            return getSubFolderIdIfExists(null, folderPath.substring(currentMailboxPath.length() + 1));
        } else if (folderPath.startsWith("/users/")) {
            int slashIndex = folderPath.indexOf('/', USERS_ROOT.length());
            String mailbox;
            String subFolderPath;
            if (slashIndex >= 0) {
                mailbox = folderPath.substring(USERS_ROOT.length(), slashIndex);
                subFolderPath = folderPath.substring(slashIndex + 1);
            } else {
                mailbox = folderPath.substring(USERS_ROOT.length());
                subFolderPath = "";
            }
            return getSubFolderIdIfExists(mailbox, subFolderPath);
        } else {
            return getSubFolderIdIfExists(null, folderPath);
        }
    }

    protected FolderId getSubFolderIdIfExists(String mailbox, String folderPath) throws IOException {
        String[] folderNames;
        FolderId currentFolderId;

        if (folderPath.startsWith(PUBLIC_ROOT)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.publicfoldersroot);
            folderNames = folderPath.substring(PUBLIC_ROOT.length()).split("/");
        } else if (folderPath.startsWith(ARCHIVE_ROOT)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.archivemsgfolderroot);
            folderNames = folderPath.substring(ARCHIVE_ROOT.length()).split("/");
        } else if (folderPath.startsWith(INBOX) || folderPath.startsWith(LOWER_CASE_INBOX)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.inbox);
            folderNames = folderPath.substring(INBOX.length()).split("/");
        } else if (folderPath.startsWith(CALENDAR)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.calendar);
            folderNames = folderPath.substring(CALENDAR.length()).split("/");
        } else if (folderPath.startsWith(TASKS)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.tasks);
            folderNames = folderPath.substring(TASKS.length()).split("/");
        } else if (folderPath.startsWith(CONTACTS)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.contacts);
            folderNames = folderPath.substring(CONTACTS.length()).split("/");
        } else if (folderPath.startsWith(SENT)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.sentitems);
            folderNames = folderPath.substring(SENT.length()).split("/");
        } else if (folderPath.startsWith(DRAFTS)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.drafts);
            folderNames = folderPath.substring(DRAFTS.length()).split("/");
        } else if (folderPath.startsWith(TRASH)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.deleteditems);
            folderNames = folderPath.substring(TRASH.length()).split("/");
        } else if (folderPath.startsWith(JUNK)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.junkemail);
            folderNames = folderPath.substring(JUNK.length()).split("/");
        } else if (folderPath.startsWith(UNSENT)) {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.outbox);
            folderNames = folderPath.substring(UNSENT.length()).split("/");
        } else {
            currentFolderId = DistinguishedFolderId.getInstance(mailbox, DistinguishedFolderId.Name.msgfolderroot);
            folderNames = folderPath.split("/");
        }
        for (String folderName : folderNames) {
            if (folderName.length() > 0) {
                currentFolderId = getSubFolderByName(currentFolderId, folderName);
                if (currentFolderId == null) {
                    break;
                }
            }
        }
        return currentFolderId;
    }

    protected FolderId getSubFolderByName(FolderId parentFolderId, String folderName) throws IOException {
        FolderId folderId = null;
        FindFolderMethod findFolderMethod = new FindFolderMethod(
                FolderQueryTraversal.SHALLOW,
                BaseShape.ID_ONLY,
                parentFolderId,
                FOLDER_PROPERTIES,
                new TwoOperandExpression(TwoOperandExpression.Operator.IsEqualTo,
                        Field.get("folderDisplayName"), folderName)
        );
        executeMethod(findFolderMethod);
        EWSMethod.Item item = findFolderMethod.getResponseItem();
        if (item != null) {
            folderId = new FolderId(item);
        }
        return folderId;
    }

    protected void executeMethod(EWSMethod ewsMethod) throws IOException {
        try {
            ewsMethod.setServerVersion(serverVersion);
            httpClient.executeMethod(ewsMethod);
            if (serverVersion == null) {
                serverVersion = ewsMethod.getServerVersion();
            }
            ewsMethod.checkSuccess();
        } finally {
            ewsMethod.releaseConnection();
        }
    }

    protected static final HashMap<String, String> GALFIND_ATTRIBUTE_MAP = new HashMap<String, String>();

    static {
        GALFIND_ATTRIBUTE_MAP.put("imapUid", "Name");
        GALFIND_ATTRIBUTE_MAP.put("cn", "DisplayName");
        GALFIND_ATTRIBUTE_MAP.put("givenName", "GivenName");
        GALFIND_ATTRIBUTE_MAP.put("sn", "Surname");
        GALFIND_ATTRIBUTE_MAP.put("smtpemail1", "EmailAddress");

        GALFIND_ATTRIBUTE_MAP.put("roomnumber", "OfficeLocation");
        GALFIND_ATTRIBUTE_MAP.put("street", "BusinessStreet");
        GALFIND_ATTRIBUTE_MAP.put("l", "BusinessCity");
        GALFIND_ATTRIBUTE_MAP.put("o", "CompanyName");
        GALFIND_ATTRIBUTE_MAP.put("postalcode", "BusinessPostalCode");
        GALFIND_ATTRIBUTE_MAP.put("st", "BusinessState");
        GALFIND_ATTRIBUTE_MAP.put("co", "BusinessCountryOrRegion");

        GALFIND_ATTRIBUTE_MAP.put("manager", "Manager");
        GALFIND_ATTRIBUTE_MAP.put("middlename", "Initials");
        GALFIND_ATTRIBUTE_MAP.put("title", "JobTitle");
        GALFIND_ATTRIBUTE_MAP.put("department", "Department");

        GALFIND_ATTRIBUTE_MAP.put("otherTelephone", "OtherTelephone");
        GALFIND_ATTRIBUTE_MAP.put("telephoneNumber", "BusinessPhone");
        GALFIND_ATTRIBUTE_MAP.put("mobile", "MobilePhone");
        GALFIND_ATTRIBUTE_MAP.put("facsimiletelephonenumber", "BusinessFax");
        GALFIND_ATTRIBUTE_MAP.put("secretarycn", "AssistantName");
    }

    protected static final HashSet<String> IGNORE_ATTRIBUTE_SET = new HashSet<String>();

    static {
        IGNORE_ATTRIBUTE_SET.add("ContactSource");
        IGNORE_ATTRIBUTE_SET.add("Culture");
        IGNORE_ATTRIBUTE_SET.add("AssistantPhone");
    }

    protected EwsContact buildGalfindContact(EWSMethod.Item response) {
        EwsContact contact = new EwsContact(this);
        contact.setName(response.get("Name"));
        contact.put("imapUid", response.get("Name"));
        contact.put("uid", response.get("Name"));
        if (LOGGER.isDebugEnabled()) {
            for (Map.Entry<String, String> entry : response.entrySet()) {
                String key = entry.getKey();
                if (!IGNORE_ATTRIBUTE_SET.contains(key) && !GALFIND_ATTRIBUTE_MAP.containsValue(key)) {
                    LOGGER.debug("Unsupported ResolveNames " + contact.getName() + " response attribute: " + key + " value: " + entry.getValue());
                }
            }
        }
        for (Map.Entry<String, String> entry : GALFIND_ATTRIBUTE_MAP.entrySet()) {
            String attributeValue = response.get(entry.getValue());
            if (attributeValue != null) {
                contact.put(entry.getKey(), attributeValue);
            }
        }
        return contact;
    }

    @Override
    public Map<String, Contact> galFind(Condition condition, Set<String> returningAttributes, int sizeLimit) throws IOException {
        Map<String, Contact> contacts = new HashMap<String, Contact>();
        if (condition instanceof MultiCondition) {
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
            String mappedAttributeName = GALFIND_ATTRIBUTE_MAP.get(((davmail.exchange.condition.AttributeCondition) condition).getAttributeName());
            if (mappedAttributeName != null) {
                String value = ((AttributeCondition) condition).getValue().toLowerCase();
                Operator operator = ((AttributeCondition) condition).getOperator();
                String searchValue = value;
                if (mappedAttributeName.startsWith("EmailAddress")) {
                    searchValue = "smtp:" + searchValue;
                }
                if (operator == Operator.IsEqualTo) {
                    searchValue = '=' + searchValue;
                }
                ResolveNamesMethod resolveNamesMethod = new ResolveNamesMethod(searchValue);
                executeMethod(resolveNamesMethod);
                List<EWSMethod.Item> responses = resolveNamesMethod.getResponseItems();
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug("ResolveNames(" + searchValue + ") returned " + responses.size() + " results");
                }
                for (EWSMethod.Item response : responses) {
                    EwsContact contact = buildGalfindContact(response);
                    if (condition.isMatch(contact)) {
                        contacts.put(contact.getName().toLowerCase(), contact);
                    }
                }
            }
        }
        return contacts;
    }

    protected Date parseDateFromExchange(String exchangeDateValue) throws DavMailException {
        Date dateValue = null;
        if (exchangeDateValue != null) {
            try {
                dateValue = getExchangeZuluDateFormat().parse(exchangeDateValue);
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", exchangeDateValue);
            }
        }
        return dateValue;
    }

    protected String convertDateFromExchange(String exchangeDateValue) throws DavMailException {
        String zuluDateValue = null;
        if (exchangeDateValue != null) {
            try {
                zuluDateValue = getZuluDateFormat().format(getExchangeZuluDateFormat().parse(exchangeDateValue));
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", exchangeDateValue);
            }
        }
        return zuluDateValue;
    }

    protected String convertCalendarDateToExchange(String vcalendarDateValue) throws DavMailException {
        String zuluDateValue = null;
        if (vcalendarDateValue != null) {
            try {
                SimpleDateFormat dateParser;
                if (vcalendarDateValue.length() == 8) {
                    dateParser = new SimpleDateFormat("yyyyMMdd", Locale.ENGLISH);
                } else {
                    dateParser = new SimpleDateFormat("yyyyMMdd'T'HHmmss", Locale.ENGLISH);
                }
                dateParser.setTimeZone(GMT_TIMEZONE);
                SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss", Locale.ENGLISH);
                dateFormatter.setTimeZone(GMT_TIMEZONE);
                zuluDateValue = dateFormatter.format(dateParser.parse(vcalendarDateValue));
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", vcalendarDateValue);
            }
        }
        return zuluDateValue;
    }

    protected String convertDateFromExchangeToTaskDate(String exchangeDateValue) throws DavMailException {
        String zuluDateValue = null;
        if (exchangeDateValue != null) {
            try {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd", Locale.ENGLISH);
                dateFormat.setTimeZone(GMT_TIMEZONE);
                zuluDateValue = dateFormat.format(getExchangeZuluDateFormat().parse(exchangeDateValue));
            } catch (ParseException e) {
                throw new DavMailException("EXCEPTION_INVALID_DATE", exchangeDateValue);
            }
        }
        return zuluDateValue;
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
                result = ExchangeSession.getExchangeZuluDateFormat().format(calendarValue.getTime());
            } catch (ParseException e) {
                LOGGER.warn("Invalid date: " + value);
            }
        }

        return result;
    }

    /**
     * Format date to exchange search format.
     *
     * @param date date object
     * @return formatted search date
     */
    @Override
    public String formatSearchDate(Date date) {
        SimpleDateFormat dateFormatter = new SimpleDateFormat(YYYY_MM_DD_T_HHMMSS_Z, Locale.ENGLISH);
        dateFormatter.setTimeZone(GMT_TIMEZONE);
        return dateFormatter.format(date);
    }

    /**
     * Check if itemName is long and base64 encoded.
     * User generated item names are usually short
     * @param itemName item name
     * @return true if itemName is an EWS item id
     */
    protected static boolean isItemId(String itemName) {
        return itemName.length() >= 152
                // item name is base64
                && itemName.matches("^([A-Za-z0-9+/]{4})*([A-Za-z0-9+/]{4}|[A-Za-z0-9+/]{3}=|[A-Za-z0-9+/]{2}==)$");
    }


    protected static final Map<String, String> importanceToPriorityMap = new HashMap<String, String>();

    static {
        importanceToPriorityMap.put("High", "1");
        importanceToPriorityMap.put("Normal", "5");
        importanceToPriorityMap.put("Low", "9");
    }

    protected static final Map<String, String> priorityToImportanceMap = new HashMap<String, String>();

    static {
        priorityToImportanceMap.put("1", "High");
        priorityToImportanceMap.put("5", "Normal");
        priorityToImportanceMap.put("9", "Low");
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
}

