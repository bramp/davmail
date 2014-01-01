package davmail.exchange.ews;

import davmail.exchange.*;
import davmail.exchange.entity.ItemResult;
import davmail.util.StringUtil;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.httpclient.HttpStatus;
import org.apache.log4j.Logger;

import javax.mail.MessagingException;
import javax.mail.internet.InternetAddress;
import javax.mail.util.SharedByteArrayInputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.util.*;

class EwsEvent extends davmail.exchange.entity.Event {
    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private final EwsExchangeSession ewsExchangeSession;

    // item id
    ItemId itemId;
    String type;
    boolean isException;

    EwsEvent(EwsExchangeSession ewsExchangeSession, EWSMethod.Item response) {
        super(ewsExchangeSession);

        this.ewsExchangeSession = ewsExchangeSession;
        itemId = new ItemId(response);

        type = response.type;

        permanentUrl = response.get(Field.get("permanenturl").getResponseName());
        etag = response.get(Field.get("etag").getResponseName());
        displayName = response.get(Field.get("displayname").getResponseName());
        subject = response.get(Field.get("subject").getResponseName());
        itemName = StringUtil.decodeUrlcompname(response.get(Field.get("urlcompname").getResponseName()));
        // workaround for missing urlcompname in Exchange 2010
        if (itemName == null) {
            itemName = StringUtil.base64ToUrl(itemId.id) + ".EML";
        }
        String instancetype = response.get(Field.get("instancetype").getResponseName());

        // TODO BUG? isrecurring and calendaritemtype are not used.
        boolean isrecurring = "true".equals(response.get(Field.get("isrecurring").getResponseName()));
        String calendaritemtype = response.get(Field.get("calendaritemtype").getResponseName());
        isException = "3".equals(instancetype);
    }

    /**
     * @inheritDoc
     */
    EwsEvent(EwsExchangeSession ewsExchangeSession, String folderPath, String itemName, String contentClass, String itemBody, String etag, String noneMatch) throws IOException {
        super(ewsExchangeSession, folderPath, itemName, contentClass, itemBody, etag, noneMatch);
        this.ewsExchangeSession = ewsExchangeSession;
    }

    @Override
    public ItemResult createOrUpdate() throws IOException {
        if (vCalendar.isTodo() && ewsExchangeSession.isMainCalendar(folderPath)) {
            // task item, move to tasks folder
            folderPath = ExchangeSession.TASKS;
        }

        ItemResult itemResult = new ItemResult();
        EWSMethod createOrUpdateItemMethod;

        // first try to load existing event
        String currentEtag = null;
        ItemId currentItemId = null;
        String ownerResponseReply = null;

        EWSMethod.Item currentItem = ewsExchangeSession.getEwsItem(folderPath, itemName);
        if (currentItem != null) {
            currentItemId = new ItemId(currentItem);
            currentEtag = currentItem.get(Field.get("etag").getResponseName());
            LOGGER.debug("Existing item found with etag: " + currentEtag + " client etag: " + etag + " id: " + currentItemId.id);
        }
        if ("*".equals(noneMatch)) {
            // create requested
            if (currentItemId != null) {
                itemResult.status = HttpStatus.SC_PRECONDITION_FAILED;
                return itemResult;
            }
        } else if (etag != null) {
            // update requested
            if (currentItemId == null || !etag.equals(currentEtag)) {
                itemResult.status = HttpStatus.SC_PRECONDITION_FAILED;
                return itemResult;
            }
        }
        if (vCalendar.isTodo()) {
            // create or update task method
            EWSMethod.Item newItem = new EWSMethod.Item();
            newItem.type = "Task";
            List<FieldUpdate> updates = new ArrayList<FieldUpdate>();
            updates.add(Field.createFieldUpdate("importance", ewsExchangeSession.convertPriorityToExchange(vCalendar.getFirstVeventPropertyValue("PRIORITY"))));
            updates.add(Field.createFieldUpdate("calendaruid", vCalendar.getFirstVeventPropertyValue("UID")));
            // force urlcompname
            updates.add(Field.createFieldUpdate("urlcompname", exchangeSession.convertItemNameToEML(itemName)));
            updates.add(Field.createFieldUpdate("subject", vCalendar.getFirstVeventPropertyValue("SUMMARY")));
            updates.add(Field.createFieldUpdate("description", vCalendar.getFirstVeventPropertyValue("DESCRIPTION")));
            updates.add(Field.createFieldUpdate("keywords", vCalendar.getFirstVeventPropertyValue("CATEGORIES")));
            updates.add(Field.createFieldUpdate("startdate", ewsExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DTSTART"))));
            updates.add(Field.createFieldUpdate("duedate", ewsExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DUE"))));
            updates.add(Field.createFieldUpdate("datecompleted", ewsExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("COMPLETED"))));

            updates.add(Field.createFieldUpdate("commonstart", ewsExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DTSTART"))));
            updates.add(Field.createFieldUpdate("commonend", ewsExchangeSession.convertTaskDateToZulu(vCalendar.getFirstVeventPropertyValue("DUE"))));

            String percentComplete = vCalendar.getFirstVeventPropertyValue("PERCENT-COMPLETE");
            if (percentComplete == null) {
                percentComplete = "0";
            }
            updates.add(Field.createFieldUpdate("percentcomplete", percentComplete));
            String vTodoStatus = vCalendar.getFirstVeventPropertyValue("STATUS");
            if (vTodoStatus == null) {
                updates.add(Field.createFieldUpdate("taskstatus", "NotStarted"));
            } else {
                updates.add(Field.createFieldUpdate("taskstatus", EwsExchangeSession.vTodoToTaskStatusMap.get(vTodoStatus)));
            }

            //updates.add(Field.createFieldUpdate("iscomplete", "COMPLETED".equals(vTodoStatus)?"True":"False"));

            if (currentItemId != null) {
                // update
                createOrUpdateItemMethod = new UpdateItemMethod(MessageDisposition.SaveOnly,
                        ConflictResolution.AutoResolve,
                        SendMeetingInvitationsOrCancellations.SendToNone,
                        currentItemId, updates);
            } else {
                newItem.setFieldUpdates(updates);
                // create
                createOrUpdateItemMethod = new CreateItemMethod(MessageDisposition.SaveOnly, SendMeetingInvitations.SendToNone, ewsExchangeSession.getFolderId(folderPath), newItem);
            }

        } else {

            if (currentItemId != null) {
                /*Set<FieldUpdate> updates = new HashSet<FieldUpdate>();
                // TODO: update properties instead of brute force delete/add
                updates.add(new FieldUpdate(Field.get("mimeContent"), new String(Base64.encodeBase64(itemContent))));
                // update
                createOrUpdateItemMethod = new UpdateItemMethod(MessageDisposition.SaveOnly,
                       ConflictResolution.AutoResolve,
                       SendMeetingInvitationsOrCancellations.SendToNone,
                       currentItemId, updates);*/
                // hard method: delete/create on update
                DeleteItemMethod deleteItemMethod = new DeleteItemMethod(currentItemId, DeleteType.HardDelete, SendMeetingCancellations.SendToNone);
                ewsExchangeSession.executeMethod(deleteItemMethod);
            } //else {
            // create
            EWSMethod.Item newItem = new EWSMethod.Item();
            newItem.type = "CalendarItem";
            newItem.mimeContent = Base64.encodeBase64(vCalendar.toString().getBytes("UTF-8"));
            ArrayList<FieldUpdate> updates = new ArrayList<FieldUpdate>();
            if (!vCalendar.hasVAlarm()) {
                updates.add(Field.createFieldUpdate("reminderset", "false"));
            }
            //updates.add(Field.createFieldUpdate("outlookmessageclass", "IPM.Appointment"));
            // force urlcompname
            updates.add(Field.createFieldUpdate("urlcompname", exchangeSession.convertItemNameToEML(itemName)));
            if (vCalendar.isMeeting()) {
                if (vCalendar.isMeetingOrganizer()) {
                    updates.add(Field.createFieldUpdate("apptstateflags", "1"));
                } else {
                    updates.add(Field.createFieldUpdate("apptstateflags", "3"));
                }
            } else {
                updates.add(Field.createFieldUpdate("apptstateflags", "0"));
            }
            // store mozilla invitations option
            String xMozSendInvitations = vCalendar.getFirstVeventPropertyValue("X-MOZ-SEND-INVITATIONS");
            if (xMozSendInvitations != null) {
                updates.add(Field.createFieldUpdate("xmozsendinvitations", xMozSendInvitations));
            }
            // handle mozilla alarm
            String xMozLastack = vCalendar.getFirstVeventPropertyValue("X-MOZ-LASTACK");
            if (xMozLastack != null) {
                updates.add(Field.createFieldUpdate("xmozlastack", xMozLastack));
            }
            String xMozSnoozeTime = vCalendar.getFirstVeventPropertyValue("X-MOZ-SNOOZE-TIME");
            if (xMozSnoozeTime != null) {
                updates.add(Field.createFieldUpdate("xmozsnoozetime", xMozSnoozeTime));
            }

            if (vCalendar.isMeeting() && exchangeSession.getServerVersion().isExchange2007()) {
                Set<String> requiredAttendees = new HashSet<String>();
                Set<String> optionalAttendees = new HashSet<String>();
                List<VProperty> attendeeProperties = vCalendar.getFirstVeventProperties("ATTENDEE");
                if (attendeeProperties != null) {
                    for (VProperty property : attendeeProperties) {
                        String attendeeEmail = vCalendar.getEmailValue(property);
                        if (attendeeEmail != null && attendeeEmail.indexOf('@') >= 0) {
                            if (ewsExchangeSession.email.equals(attendeeEmail)) {
                                String ownerPartStat = property.getParamValue("PARTSTAT");
                                if ("ACCEPTED".equals(ownerPartStat)) {
                                    ownerResponseReply = "AcceptItem";
                                    // do not send DeclineItem to avoid deleting target event
                                    //} else if  ("DECLINED".equals(ownerPartStat)) {
                                    //    ownerResponseReply = "DeclineItem";
                                } else if ("TENTATIVE".equals(ownerPartStat)) {
                                    ownerResponseReply = "TentativelyAcceptItem";
                                }
                            }
                            InternetAddress internetAddress = new InternetAddress(attendeeEmail, property.getParamValue("CN"));
                            String attendeeRole = property.getParamValue("ROLE");
                            if ("REQ-PARTICIPANT".equals(attendeeRole)) {
                                requiredAttendees.add(internetAddress.toString());
                            } else {
                                optionalAttendees.add(internetAddress.toString());
                            }
                        }
                    }
                }
                List<VProperty> organizerProperties = vCalendar.getFirstVeventProperties("ORGANIZER");
                if (organizerProperties != null) {
                    VProperty property = organizerProperties.get(0);
                    String organizerEmail = vCalendar.getEmailValue(property);
                    if (organizerEmail != null && organizerEmail.indexOf('@') >= 0) {
                        updates.add(Field.createFieldUpdate("from", organizerEmail));
                    }
                }

                if (requiredAttendees.size() > 0) {
                    updates.add(Field.createFieldUpdate("to", StringUtil.join(requiredAttendees, ", ")));
                }
                if (optionalAttendees.size() > 0) {
                    updates.add(Field.createFieldUpdate("cc", StringUtil.join(optionalAttendees, ", ")));
                }
            }

            // patch allday date values, only on 2007
            boolean isAllDayExchange2007 = exchangeSession.getServerVersion().isExchange2007() && vCalendar.isCdoAllDay();
            if (isAllDayExchange2007) {
                updates.add(Field.createFieldUpdate("dtstart", ewsExchangeSession.convertCalendarDateToExchange(vCalendar.getFirstVeventPropertyValue("DTSTART"))));
                updates.add(Field.createFieldUpdate("dtend", ewsExchangeSession.convertCalendarDateToExchange(vCalendar.getFirstVeventPropertyValue("DTEND"))));
            }
            updates.add(Field.createFieldUpdate("busystatus", "BUSY".equals(vCalendar.getFirstVeventPropertyValue("X-MICROSOFT-CDO-BUSYSTATUS")) ? "Busy" : "Free"));
            if (isAllDayExchange2007) {
                updates.add(Field.createFieldUpdate("meetingtimezone", vCalendar.getVTimezone().getPropertyValue("TZID")));
            }

            newItem.setFieldUpdates(updates);
            createOrUpdateItemMethod = new CreateItemMethod(MessageDisposition.SaveOnly, SendMeetingInvitations.SendToNone, ewsExchangeSession.getFolderId(folderPath), newItem);
            // force context Timezone on Exchange 2010
            if (exchangeSession.getServerVersion().isExchange2010()) {
                createOrUpdateItemMethod.setTimezoneContext(ewsExchangeSession.getVTimezone().getPropertyValue("TZID"));
            }
            //}
        }
        ewsExchangeSession.executeMethod(createOrUpdateItemMethod);

        itemResult.status = createOrUpdateItemMethod.getStatusCode();
        if (itemResult.status == HttpURLConnection.HTTP_OK) {
            //noinspection VariableNotUsedInsideIf
            if (currentItemId == null) {
                itemResult.status = HttpStatus.SC_CREATED;
                LOGGER.debug("Created event " + getHref());
            } else {
                LOGGER.warn("Overwritten event " + getHref());
            }
        }

        // force responsetype on Exchange 2007
        if (ownerResponseReply != null) {
            EWSMethod.Item responseTypeItem = new EWSMethod.Item();
            responseTypeItem.referenceItemId = new ItemId("ReferenceItemId", createOrUpdateItemMethod.getResponseItem());
            responseTypeItem.type = ownerResponseReply;
            createOrUpdateItemMethod = new CreateItemMethod(MessageDisposition.SaveOnly, SendMeetingInvitations.SendToNone, null, responseTypeItem);
            ewsExchangeSession.executeMethod(createOrUpdateItemMethod);

            // force urlcompname again
            ArrayList<FieldUpdate> updates = new ArrayList<FieldUpdate>();
            updates.add(Field.createFieldUpdate("urlcompname", exchangeSession.convertItemNameToEML(itemName)));
            createOrUpdateItemMethod = new UpdateItemMethod(MessageDisposition.SaveOnly,
                    ConflictResolution.AlwaysOverwrite,
                    SendMeetingInvitationsOrCancellations.SendToNone,
                    new ItemId(createOrUpdateItemMethod.getResponseItem()),
                    updates);
            ewsExchangeSession.executeMethod(createOrUpdateItemMethod);
        }

        ItemId newItemId = new ItemId(createOrUpdateItemMethod.getResponseItem());
        GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, newItemId, false);
        getItemMethod.addAdditionalProperty(Field.get("etag"));
        ewsExchangeSession.executeMethod(getItemMethod);
        itemResult.etag = getItemMethod.getResponseItem().get(Field.get("etag").getResponseName());

        return itemResult;

    }

    @Override
    public byte[] getEventContent() throws IOException {
        byte[] content;
        LOGGER.debug("Get event: " + itemName);

        try {
            GetItemMethod getItemMethod;
            if ("Task".equals(type)) {
                getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, false);
                getItemMethod.addAdditionalProperty(Field.get("importance"));
                getItemMethod.addAdditionalProperty(Field.get("subject"));
                getItemMethod.addAdditionalProperty(Field.get("created"));
                getItemMethod.addAdditionalProperty(Field.get("lastmodified"));
                getItemMethod.addAdditionalProperty(Field.get("calendaruid"));
                getItemMethod.addAdditionalProperty(Field.get("description"));
                getItemMethod.addAdditionalProperty(Field.get("percentcomplete"));
                getItemMethod.addAdditionalProperty(Field.get("taskstatus"));
                getItemMethod.addAdditionalProperty(Field.get("startdate"));
                getItemMethod.addAdditionalProperty(Field.get("duedate"));
                getItemMethod.addAdditionalProperty(Field.get("datecompleted"));
                getItemMethod.addAdditionalProperty(Field.get("keywords"));

            } else if (!"Message".equals(type)) {
                getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, true);
                getItemMethod.addAdditionalProperty(Field.get("reminderset"));
                getItemMethod.addAdditionalProperty(Field.get("calendaruid"));
                getItemMethod.addAdditionalProperty(Field.get("myresponsetype"));
                getItemMethod.addAdditionalProperty(Field.get("requiredattendees"));
                getItemMethod.addAdditionalProperty(Field.get("optionalattendees"));
                getItemMethod.addAdditionalProperty(Field.get("modifiedoccurrences"));
                getItemMethod.addAdditionalProperty(Field.get("xmozlastack"));
                getItemMethod.addAdditionalProperty(Field.get("xmozsnoozetime"));
                getItemMethod.addAdditionalProperty(Field.get("xmozsendinvitations"));
            } else {
                getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, true);
            }

            ewsExchangeSession.executeMethod(getItemMethod);
            if ("Task".equals(type)) {
                VObject vTimezone = ewsExchangeSession.getVTimezone();
                VCalendar localVCalendar = new VCalendar();
                VObject vTodo = new VObject();
                vTodo.type = "VTODO";
                localVCalendar.setTimezone(vTimezone);
                vTodo.setPropertyValue("LAST-MODIFIED", ewsExchangeSession.convertDateFromExchange(getItemMethod.getResponseItem().get(Field.get("lastmodified").getResponseName())));
                vTodo.setPropertyValue("CREATED", ewsExchangeSession.convertDateFromExchange(getItemMethod.getResponseItem().get(Field.get("created").getResponseName())));
                String calendarUid = getItemMethod.getResponseItem().get(Field.get("calendaruid").getResponseName());
                if (calendarUid == null) {
                    // use item id as uid for Exchange created tasks
                    calendarUid = itemId.id;
                }
                vTodo.setPropertyValue("UID", calendarUid);
                vTodo.setPropertyValue("SUMMARY", getItemMethod.getResponseItem().get(Field.get("subject").getResponseName()));
                vTodo.setPropertyValue("DESCRIPTION", getItemMethod.getResponseItem().get(Field.get("description").getResponseName()));
                vTodo.setPropertyValue("PRIORITY", ewsExchangeSession.convertPriorityFromExchange(getItemMethod.getResponseItem().get(Field.get("importance").getResponseName())));
                vTodo.setPropertyValue("PERCENT-COMPLETE", getItemMethod.getResponseItem().get(Field.get("percentcomplete").getResponseName()));
                vTodo.setPropertyValue("STATUS", EwsExchangeSession.taskTovTodoStatusMap.get(getItemMethod.getResponseItem().get(Field.get("taskstatus").getResponseName())));

                vTodo.setPropertyValue("DUE;VALUE=DATE", ewsExchangeSession.convertDateFromExchangeToTaskDate(getItemMethod.getResponseItem().get(Field.get("duedate").getResponseName())));
                vTodo.setPropertyValue("DTSTART;VALUE=DATE", ewsExchangeSession.convertDateFromExchangeToTaskDate(getItemMethod.getResponseItem().get(Field.get("startdate").getResponseName())));
                vTodo.setPropertyValue("COMPLETED;VALUE=DATE", ewsExchangeSession.convertDateFromExchangeToTaskDate(getItemMethod.getResponseItem().get(Field.get("datecompleted").getResponseName())));

                vTodo.setPropertyValue("CATEGORIES", getItemMethod.getResponseItem().get(Field.get("keywords").getResponseName()));

                localVCalendar.addVObject(vTodo);
                content = localVCalendar.toString().getBytes("UTF-8");
            } else {
                content = getItemMethod.getMimeContent();
                if (content == null) {
                    throw new IOException("empty event body");
                }
                if (!"CalendarItem".equals(type)) {
                    content = getICS(new SharedByteArrayInputStream(content));
                }
                VCalendar localVCalendar = new VCalendar(content, ewsExchangeSession.email, ewsExchangeSession.getVTimezone());

                String calendaruid = getItemMethod.getResponseItem().get(Field.get("calendaruid").getResponseName());

                if (exchangeSession.getServerVersion().isExchange2007()) {
                    // remove additional reminder
                    if (!"true".equals(getItemMethod.getResponseItem().get(Field.get("reminderset").getResponseName()))) {
                        localVCalendar.removeVAlarm();
                    }
                    if (calendaruid != null) {
                        localVCalendar.setFirstVeventPropertyValue("UID", calendaruid);
                    }
                }
                fixAttendees(getItemMethod, localVCalendar.getFirstVevent());
                // fix UID and RECURRENCE-ID, broken at least on Exchange 2007
                List<EWSMethod.Occurrence> occurences = getItemMethod.getResponseItem().getOccurrences();
                if (occurences != null) {
                    Iterator<VObject> modifiedOccurrencesIterator = localVCalendar.getModifiedOccurrences().iterator();
                    for (EWSMethod.Occurrence occurrence : occurences) {
                        if (modifiedOccurrencesIterator.hasNext()) {
                            VObject modifiedOccurrence = modifiedOccurrencesIterator.next();
                            // fix modified occurrences attendees
                            GetItemMethod getOccurrenceMethod = new GetItemMethod(BaseShape.ID_ONLY, occurrence.itemId, false);
                            getOccurrenceMethod.addAdditionalProperty(Field.get("requiredattendees"));
                            getOccurrenceMethod.addAdditionalProperty(Field.get("optionalattendees"));
                            getOccurrenceMethod.addAdditionalProperty(Field.get("modifiedoccurrences"));
                            ewsExchangeSession.executeMethod(getOccurrenceMethod);
                            fixAttendees(getOccurrenceMethod, modifiedOccurrence);

                            if (exchangeSession.getServerVersion().isExchange2007()) {
                                // fix uid, should be the same as main VEVENT
                                if (calendaruid != null) {
                                    modifiedOccurrence.setPropertyValue("UID", calendaruid);
                                }

                                VProperty recurrenceId = modifiedOccurrence.getProperty("RECURRENCE-ID");
                                if (recurrenceId != null) {
                                    recurrenceId.removeParam("TZID");
                                    recurrenceId.getValues().set(0, ewsExchangeSession.convertDateFromExchange(occurrence.originalStart));
                                }
                            }
                        }
                    }
                }
                // restore mozilla invitations option
                localVCalendar.setFirstVeventPropertyValue("X-MOZ-SEND-INVITATIONS",
                        getItemMethod.getResponseItem().get(Field.get("xmozsendinvitations").getResponseName()));
                // restore mozilla alarm status
                localVCalendar.setFirstVeventPropertyValue("X-MOZ-LASTACK",
                        getItemMethod.getResponseItem().get(Field.get("xmozlastack").getResponseName()));
                localVCalendar.setFirstVeventPropertyValue("X-MOZ-SNOOZE-TIME",
                        getItemMethod.getResponseItem().get(Field.get("xmozsnoozetime").getResponseName()));
                // overwrite method
                // localVCalendar.setPropertyValue("METHOD", "REQUEST");
                content = localVCalendar.toString().getBytes("UTF-8");
            }
        } catch (IOException e) {
            throw buildHttpException(e);
        } catch (MessagingException e) {
            throw buildHttpException(e);
        }
        return content;
    }

    protected void fixAttendees(GetItemMethod getItemMethod, VObject vEvent) throws EWSException {
        List<EWSMethod.Attendee> attendees = getItemMethod.getResponseItem().getAttendees();
        if (attendees != null) {
            for (EWSMethod.Attendee attendee : attendees) {
                VProperty attendeeProperty = new VProperty("ATTENDEE", "mailto:" + attendee.email);
                attendeeProperty.addParam("CN", attendee.name);
                String myResponseType = getItemMethod.getResponseItem().get(Field.get("myresponsetype").getResponseName());
                if (exchangeSession.getServerVersion().isExchange2007() && ewsExchangeSession.email.equalsIgnoreCase(attendee.email) && myResponseType != null) {
                    attendeeProperty.addParam("PARTSTAT", EWSMethod.responseTypeToPartstat(myResponseType));
                } else {
                    attendeeProperty.addParam("PARTSTAT", attendee.partstat);
                }
                //attendeeProperty.addParam("RSVP", "TRUE");
                attendeeProperty.addParam("ROLE", attendee.role);
                vEvent.addProperty(attendeeProperty);
            }
        }
    }
}
