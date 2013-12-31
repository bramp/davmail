package davmail.exchange.dav;

import davmail.Settings;
import davmail.exception.HttpNotFoundException;
import davmail.exchange.VCalendar;
import davmail.exchange.VObject;
import davmail.exchange.VProperty;
import davmail.exchange.entity.Item;
import davmail.exchange.entity.ItemResult;
import davmail.http.DavGatewayHttpClientFacade;
import davmail.util.StringUtil;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.httpclient.HttpException;
import org.apache.commons.httpclient.HttpStatus;
import org.apache.commons.httpclient.URIException;
import org.apache.commons.httpclient.methods.ByteArrayRequestEntity;
import org.apache.commons.httpclient.methods.GetMethod;
import org.apache.commons.httpclient.methods.PutMethod;
import org.apache.commons.httpclient.util.URIUtil;
import org.apache.jackrabbit.webdav.DavException;
import org.apache.jackrabbit.webdav.MultiStatusResponse;
import org.apache.jackrabbit.webdav.client.methods.PropPatchMethod;
import org.apache.jackrabbit.webdav.property.DavPropertySet;
import org.apache.jackrabbit.webdav.property.PropEntry;
import org.apache.log4j.Logger;

import javax.mail.MessagingException;
import javax.mail.internet.InternetAddress;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

/**
 * @inheritDoc
 */
public class DavEvent extends davmail.exchange.entity.Event {
    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private final DavExchangeSession davExchangeSession;

    protected String instancetype;

    /**
     * Build Event instance from response info.
     *
     * @param multiStatusResponse response
     * @throws org.apache.commons.httpclient.URIException on error
     */
    public DavEvent(DavExchangeSession davExchangeSession, MultiStatusResponse multiStatusResponse) throws URIException {
        super(davExchangeSession);
        this.davExchangeSession = davExchangeSession;
        setHref(URIUtil.decode(multiStatusResponse.getHref()));
        DavPropertySet properties = multiStatusResponse.getProperties(HttpStatus.SC_OK);
        permanentUrl = davExchangeSession.getURLPropertyIfExists(properties, "permanenturl");
        etag = davExchangeSession.getPropertyIfExists(properties, "etag");
        displayName = davExchangeSession.getPropertyIfExists(properties, "displayname");
        subject = davExchangeSession.getPropertyIfExists(properties, "subject");
        instancetype = davExchangeSession.getPropertyIfExists(properties, "instancetype");
        contentClass = davExchangeSession.getPropertyIfExists(properties, "contentclass");
    }

    protected String getPermanentUrl() {
        return permanentUrl;
    }


    /**
     * @inheritDoc
     */
    public DavEvent(DavExchangeSession davExchangeSession, String folderPath, String itemName, String contentClass, String itemBody, String etag, String noneMatch) throws IOException {
        super(davExchangeSession, folderPath, itemName, contentClass, itemBody, etag, noneMatch);
        this.davExchangeSession = davExchangeSession;
    }

    protected byte[] getICSFromInternetContentProperty() throws IOException, DavException, MessagingException {
        byte[] result = null;
        // PropFind PR_INTERNET_CONTENT
        String propertyValue = davExchangeSession.getItemProperty(permanentUrl, "internetContent");
        if (propertyValue != null) {
            byte[] byteArray = Base64.decodeBase64(propertyValue.getBytes());
            result = getICS(new ByteArrayInputStream(byteArray));
        }
        return result;
    }

    /**
     * Load ICS content from Exchange server.
     * User Translate: f header to get MIME event content and get ICS attachment from it
     *
     * @return ICS (iCalendar) event
     * @throws org.apache.commons.httpclient.HttpException on error
     */
    @Override
    public byte[] getEventContent() throws IOException {
        byte[] result = null;
        LOGGER.debug("Get event subject: " + getSubject() + " contentclass: "+contentClass+" href: " + getHref() + " permanentUrl: " + permanentUrl);
        // do not try to load tasks MIME body
        if (!"urn:content-classes:task".equals(contentClass)) {
            // try to get PR_INTERNET_CONTENT
            try {
                result = getICSFromInternetContentProperty();
                if (result == null) {
                    GetMethod method = new GetMethod(davExchangeSession.encodeAndFixUrl(permanentUrl));
                    method.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                    method.setRequestHeader("Translate", "f");
                    try {
                        DavGatewayHttpClientFacade.executeGetMethod(exchangeSession.getHttpClient(), method, true);
                        result = getICS(method.getResponseBodyAsStream());
                    } finally {
                        method.releaseConnection();
                    }
                }
            } catch (DavException e) {
                LOGGER.warn(e.getMessage());
            } catch (IOException e) {
                LOGGER.warn(e.getMessage());
            } catch (MessagingException e) {
                LOGGER.warn(e.getMessage());
            }
        }

        // failover: rebuild event from MAPI properties
        if (result == null) {
            try {
                result = getICSFromItemProperties();
            } catch (HttpException e) {
                deleteBroken();
                throw e;
            }
        }
        // debug code
        /*if (new String(result).indexOf("VTODO") < 0) {
            LOGGER.debug("Original body: " + new String(result));
            result = getICSFromItemProperties();
            LOGGER.debug("Rebuilt body: " + new String(result));
        }*/

        return result;
    }

    private byte[] getICSFromItemProperties() throws IOException {
        byte[] result;

        // experimental: build VCALENDAR from properties

        try {
            //MultiStatusResponse[] responses = DavGatewayHttpClientFacade.executeMethod(httpClient, propFindMethod);
            Set<String> eventProperties = new HashSet<String>();
            eventProperties.add("method");

            eventProperties.add("created");
            eventProperties.add("calendarlastmodified");
            eventProperties.add("dtstamp");
            eventProperties.add("calendaruid");
            eventProperties.add("subject");
            eventProperties.add("dtstart");
            eventProperties.add("dtend");
            eventProperties.add("transparent");
            eventProperties.add("organizer");
            eventProperties.add("to");
            eventProperties.add("description");
            eventProperties.add("rrule");
            eventProperties.add("exdate");
            eventProperties.add("sensitivity");
            eventProperties.add("alldayevent");
            eventProperties.add("busystatus");
            eventProperties.add("reminderset");
            eventProperties.add("reminderdelta");
            // task
            eventProperties.add("importance");
            eventProperties.add("uid");
            eventProperties.add("taskstatus");
            eventProperties.add("percentcomplete");
            eventProperties.add("keywords");
            eventProperties.add("startdate");
            eventProperties.add("duedate");
            eventProperties.add("datecompleted");

            MultiStatusResponse[] responses = davExchangeSession.searchItems(folderPath, eventProperties, exchangeSession.isEqualTo("urlcompname", exchangeSession.convertItemNameToEML(itemName)), DavExchangeSession.FolderQueryTraversal.Shallow, 1);
            if (responses.length == 0) {
                throw new HttpNotFoundException(permanentUrl + " not found");
            }
            DavPropertySet davPropertySet = responses[0].getProperties(HttpStatus.SC_OK);
            VCalendar localVCalendar = new VCalendar();
            localVCalendar.setPropertyValue("PRODID", "-//davmail.sf.net/NONSGML DavMail Calendar V1.1//EN");
            localVCalendar.setPropertyValue("VERSION", "2.0");
            localVCalendar.setPropertyValue("METHOD", davExchangeSession.getPropertyIfExists(davPropertySet, "method"));
            VObject vEvent = new VObject();
            vEvent.setPropertyValue("CREATED", davExchangeSession.convertDateFromExchange(davExchangeSession.getPropertyIfExists(davPropertySet, "created")));
            vEvent.setPropertyValue("LAST-MODIFIED", davExchangeSession.convertDateFromExchange(davExchangeSession.getPropertyIfExists(davPropertySet, "calendarlastmodified")));
            vEvent.setPropertyValue("DTSTAMP", davExchangeSession.convertDateFromExchange(davExchangeSession.getPropertyIfExists(davPropertySet, "dtstamp")));

            String uid = davExchangeSession.getPropertyIfExists(davPropertySet, "calendaruid");
            if (uid == null) {
                uid = davExchangeSession.getPropertyIfExists(davPropertySet, "uid");
            }
            vEvent.setPropertyValue("UID", uid);
            vEvent.setPropertyValue("SUMMARY", davExchangeSession.getPropertyIfExists(davPropertySet, "subject"));
            vEvent.setPropertyValue("DESCRIPTION", davExchangeSession.getPropertyIfExists(davPropertySet, "description"));
            vEvent.setPropertyValue("PRIORITY", davExchangeSession.convertPriorityFromExchange(davExchangeSession.getPropertyIfExists(davPropertySet, "importance")));
            vEvent.setPropertyValue("CATEGORIES", davExchangeSession.getPropertyIfExists(davPropertySet, "keywords"));
            String sensitivity = davExchangeSession.getPropertyIfExists(davPropertySet, "sensitivity");
            if ("2".equals(sensitivity)) {
                vEvent.setPropertyValue("CLASS", "PRIVATE");
            } else if ("3".equals(sensitivity)) {
                vEvent.setPropertyValue("CLASS", "CONFIDENTIAL");
            } else if ("0".equals(sensitivity)) {
                vEvent.setPropertyValue("CLASS", "PUBLIC");
            }

            if (instancetype == null) {
                vEvent.type = "VTODO";
                double percentComplete = davExchangeSession.getDoublePropertyIfExists(davPropertySet, "percentcomplete");
                if (percentComplete > 0) {
                    vEvent.setPropertyValue("PERCENT-COMPLETE", String.valueOf((int) (percentComplete * 100)));
                }
                vEvent.setPropertyValue("STATUS", DavExchangeSession.taskTovTodoStatusMap.get(davExchangeSession.getPropertyIfExists(davPropertySet, "taskstatus")));
                vEvent.setPropertyValue("DUE;VALUE=DATE", davExchangeSession.convertDateFromExchangeToTaskDate(davExchangeSession.getPropertyIfExists(davPropertySet, "duedate")));
                vEvent.setPropertyValue("DTSTART;VALUE=DATE", davExchangeSession.convertDateFromExchangeToTaskDate(davExchangeSession.getPropertyIfExists(davPropertySet, "startdate")));
                vEvent.setPropertyValue("COMPLETED;VALUE=DATE", davExchangeSession.convertDateFromExchangeToTaskDate(davExchangeSession.getPropertyIfExists(davPropertySet, "datecompleted")));

            } else {
                vEvent.type = "VEVENT";
                // check mandatory dtstart value
                String dtstart = davExchangeSession.getPropertyIfExists(davPropertySet, "dtstart");
                if (dtstart != null) {
                    vEvent.setPropertyValue("DTSTART", davExchangeSession.convertDateFromExchange(dtstart));
                } else {
                    LOGGER.warn("missing dtstart on item, using fake value. Set davmail.deleteBroken=true to delete broken events");
                    vEvent.setPropertyValue("DTSTART", "20000101T000000Z");
                    deleteBroken();
                }
                // same on DTEND
                String dtend = davExchangeSession.getPropertyIfExists(davPropertySet, "dtend");
                if (dtend != null) {
                    vEvent.setPropertyValue("DTEND", davExchangeSession.convertDateFromExchange(dtend));
                } else {
                    LOGGER.warn("missing dtend on item, using fake value. Set davmail.deleteBroken=true to delete broken events");
                    vEvent.setPropertyValue("DTEND", "20000101T010000Z");
                    deleteBroken();
                }
                vEvent.setPropertyValue("TRANSP", davExchangeSession.getPropertyIfExists(davPropertySet, "transparent"));
                vEvent.setPropertyValue("RRULE", davExchangeSession.getPropertyIfExists(davPropertySet, "rrule"));
                String exdates = davExchangeSession.getPropertyIfExists(davPropertySet, "exdate");
                if (exdates != null) {
                    String[] exdatearray = exdates.split(",");
                    for (String exdate : exdatearray) {
                        vEvent.addPropertyValue("EXDATE",
                                StringUtil.convertZuluDateTimeToAllDay(davExchangeSession.convertDateFromExchange(exdate)));
                    }
                }
                String organizer = davExchangeSession.getPropertyIfExists(davPropertySet, "organizer");
                String organizerEmail = null;
                if (organizer != null) {
                    InternetAddress organizerAddress = new InternetAddress(organizer);
                    organizerEmail = organizerAddress.getAddress();
                    vEvent.setPropertyValue("ORGANIZER", "MAILTO:" + organizerEmail);
                }

                // Parse attendee list
                String toHeader = davExchangeSession.getPropertyIfExists(davPropertySet, "to");
                if (toHeader != null && !toHeader.equals(organizerEmail)) {
                    InternetAddress[] attendees = InternetAddress.parseHeader(toHeader, false);
                    for (InternetAddress attendee : attendees) {
                        if (!attendee.getAddress().equalsIgnoreCase(organizerEmail)) {
                            VProperty vProperty = new VProperty("ATTENDEE", attendee.getAddress());
                            if (attendee.getPersonal() != null) {
                                vProperty.addParam("CN", attendee.getPersonal());
                            }
                            vEvent.addProperty(vProperty);
                        }
                    }

                }
                vEvent.setPropertyValue("X-MICROSOFT-CDO-ALLDAYEVENT",
                        "1".equals(davExchangeSession.getPropertyIfExists(davPropertySet, "alldayevent")) ? "TRUE" : "FALSE");
                vEvent.setPropertyValue("X-MICROSOFT-CDO-BUSYSTATUS", davExchangeSession.getPropertyIfExists(davPropertySet, "busystatus"));

                if ("1".equals(davExchangeSession.getPropertyIfExists(davPropertySet, "reminderset"))) {
                    VObject vAlarm = new VObject();
                    vAlarm.type = "VALARM";
                    vAlarm.setPropertyValue("ACTION", "DISPLAY");
                    vAlarm.setPropertyValue("DISPLAY", "Reminder");
                    String reminderdelta = davExchangeSession.getPropertyIfExists(davPropertySet, "reminderdelta");
                    VProperty vProperty = new VProperty("TRIGGER", "-PT" + reminderdelta + 'M');
                    vProperty.addParam("VALUE", "DURATION");
                    vAlarm.addProperty(vProperty);
                    vEvent.addVObject(vAlarm);
                }
            }

            localVCalendar.addVObject(vEvent);
            result = localVCalendar.toString().getBytes("UTF-8");
        } catch (MessagingException e) {
            LOGGER.warn("Unable to rebuild event content: " + e.getMessage(), e);
            throw buildHttpException(e);
        } catch (IOException e) {
            LOGGER.warn("Unable to rebuild event content: " + e.getMessage(), e);
            throw buildHttpException(e);
        }

        return result;
    }

    protected void deleteBroken() {
        // try to delete broken event
        if (Settings.getBooleanProperty("davmail.deleteBroken")) {
            LOGGER.warn("Deleting broken event at: " + permanentUrl);
            try {
                DavGatewayHttpClientFacade.executeDeleteMethod(exchangeSession.getHttpClient(), davExchangeSession.encodeAndFixUrl(permanentUrl));
            } catch (IOException ioe) {
                LOGGER.warn("Unable to delete broken event at: " + permanentUrl);
            }
        }
    }

    protected PutMethod internalCreateOrUpdate(String encodedHref, byte[] mimeContent) throws IOException {
        PutMethod putmethod = new PutMethod(encodedHref);
        putmethod.setRequestHeader("Translate", "f");
        putmethod.setRequestHeader("Overwrite", "f");
        if (etag != null) {
            putmethod.setRequestHeader("If-Match", etag);
        }
        if (noneMatch != null) {
            putmethod.setRequestHeader("If-None-Match", noneMatch);
        }
        putmethod.setRequestHeader("Content-Type", "message/rfc822");
        putmethod.setRequestEntity(new ByteArrayRequestEntity(mimeContent, "message/rfc822"));
        try {
            exchangeSession.getHttpClient().executeMethod(putmethod);
        } finally {
            putmethod.releaseConnection();
        }
        return putmethod;
    }

    /**
     * @inheritDoc
     */
    @Override
    public ItemResult createOrUpdate() throws IOException {
        ItemResult itemResult = new ItemResult();
        if (vCalendar.isTodo()) {
            if ((davExchangeSession.mailPath + davExchangeSession.calendarName).equals(folderPath)) {
                folderPath = davExchangeSession.mailPath + davExchangeSession.tasksName;
            }
            String encodedHref = URIUtil.encodePath(getHref());
            Set<PropertyValue> propertyValues = new HashSet<PropertyValue>();
            // set contentclass on create
            if (noneMatch != null) {
                propertyValues.add(Field.createPropertyValue("contentclass", "urn:content-classes:task"));
                propertyValues.add(Field.createPropertyValue("outlookmessageclass", "IPM.Task"));
                propertyValues.add(Field.createPropertyValue("calendaruid", vCalendar.getFirstVeventPropertyValue("UID")));
            }
            propertyValues.add(Field.createPropertyValue("subject", vCalendar.getFirstVeventPropertyValue("SUMMARY")));
            propertyValues.add(Field.createPropertyValue("description", vCalendar.getFirstVeventPropertyValue("DESCRIPTION")));
            propertyValues.add(Field.createPropertyValue("importance", davExchangeSession.convertPriorityToExchange(vCalendar.getFirstVeventPropertyValue("PRIORITY"))));
            String percentComplete = vCalendar.getFirstVeventPropertyValue("PERCENT-COMPLETE");
            if (percentComplete == null) {
                percentComplete = "0";
            }
            propertyValues.add(Field.createPropertyValue("percentcomplete", String.valueOf(Double.parseDouble(percentComplete) / 100)));
            String taskStatus = DavExchangeSession.vTodoToTaskStatusMap.get(vCalendar.getFirstVeventPropertyValue("STATUS"));
            propertyValues.add(Field.createPropertyValue("taskstatus", taskStatus));
            propertyValues.add(Field.createPropertyValue("keywords", vCalendar.getFirstVeventPropertyValue("CATEGORIES")));
            propertyValues.add(Field.createPropertyValue("startdate", davExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DTSTART"))));
            propertyValues.add(Field.createPropertyValue("duedate", davExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DUE"))));
            propertyValues.add(Field.createPropertyValue("datecompleted", davExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("COMPLETED"))));

            propertyValues.add(Field.createPropertyValue("iscomplete", "2".equals(taskStatus) ? "true" : "false"));
            propertyValues.add(Field.createPropertyValue("commonstart", davExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DTSTART"))));
            propertyValues.add(Field.createPropertyValue("commonend", davExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DUE"))));

            ExchangePropPatchMethod propPatchMethod = new ExchangePropPatchMethod(encodedHref, propertyValues);
            propPatchMethod.setRequestHeader("Translate", "f");
            if (etag != null) {
                propPatchMethod.setRequestHeader("If-Match", etag);
            }
            if (noneMatch != null) {
                propPatchMethod.setRequestHeader("If-None-Match", noneMatch);
            }
            try {
                int status = exchangeSession.getHttpClient().executeMethod(propPatchMethod);

                if (status == HttpStatus.SC_MULTI_STATUS) {
                    Item newItem = davExchangeSession.getItem(folderPath, itemName);
                    itemResult.status = propPatchMethod.getResponseStatusCode();
                    itemResult.etag = newItem.etag;
                } else {
                    itemResult.status = status;
                }
            } finally {
                propPatchMethod.releaseConnection();
            }

        } else {
            String encodedHref = URIUtil.encodePath(getHref());
            byte[] mimeContent = createMimeContent();
            PutMethod putMethod = internalCreateOrUpdate(encodedHref, mimeContent);
            int status = putMethod.getStatusCode();

            if (status == HttpStatus.SC_OK) {
                LOGGER.debug("Updated event " + encodedHref);
            } else if (status == HttpStatus.SC_CREATED) {
                LOGGER.debug("Created event " + encodedHref);
            } else if (status == HttpStatus.SC_NOT_FOUND) {
                LOGGER.debug("Event not found at " + encodedHref + ", searching permanenturl by urlcompname");
                // failover, search item by urlcompname
                MultiStatusResponse[] responses = davExchangeSession.searchItems(folderPath, DavExchangeSession.EVENT_REQUEST_PROPERTIES, exchangeSession.isEqualTo("urlcompname", exchangeSession.convertItemNameToEML(itemName)), DavExchangeSession.FolderQueryTraversal.Shallow, 1);
                if (responses.length == 1) {
                    encodedHref = davExchangeSession.getPropertyIfExists(responses[0].getProperties(HttpStatus.SC_OK), "permanenturl");
                    LOGGER.warn("Event found, permanenturl is " + encodedHref);
                    putMethod = internalCreateOrUpdate(encodedHref, mimeContent);
                    status = putMethod.getStatusCode();
                    if (status == HttpStatus.SC_OK) {
                        LOGGER.debug("Updated event " + encodedHref);
                    } else {
                        LOGGER.warn("Unable to create or update event " + status + ' ' + putMethod.getStatusLine());
                    }
                }
            } else {
                LOGGER.warn("Unable to create or update event " + status + ' ' + putMethod.getStatusLine());
            }

            // 440 means forbidden on Exchange
            if (status == 440) {
                status = HttpStatus.SC_FORBIDDEN;
            } else if (status == HttpStatus.SC_UNAUTHORIZED && getHref().startsWith("/public")) {
                LOGGER.warn("Ignore 401 unauthorized on public event");
                status = HttpStatus.SC_OK;
            }
            itemResult.status = status;
            if (putMethod.getResponseHeader("GetETag") != null) {
                itemResult.etag = putMethod.getResponseHeader("GetETag").getValue();
            }

            // trigger activeSync push event, only if davmail.forceActiveSyncUpdate setting is true
            if ((status == HttpStatus.SC_OK || status == HttpStatus.SC_CREATED) &&
                    (Settings.getBooleanProperty("davmail.forceActiveSyncUpdate"))) {
                ArrayList<PropEntry> propertyList = new ArrayList<PropEntry>();
                // Set contentclass to make ActiveSync happy
                propertyList.add(Field.createDavProperty("contentclass", contentClass));
                // ... but also set PR_INTERNET_CONTENT to preserve custom properties
                propertyList.add(Field.createDavProperty("internetContent", new String(Base64.encodeBase64(mimeContent))));
                PropPatchMethod propPatchMethod = new PropPatchMethod(encodedHref, propertyList);
                int patchStatus = DavGatewayHttpClientFacade.executeHttpMethod(exchangeSession.getHttpClient(), propPatchMethod);
                if (patchStatus != HttpStatus.SC_MULTI_STATUS) {
                    LOGGER.warn("Unable to patch event to trigger activeSync push");
                } else {
                    // need to retrieve new etag
                    Item newItem = exchangeSession.getItem(folderPath, itemName);
                    itemResult.etag = newItem.etag;
                }
            }
        }
        return itemResult;
    }
}
