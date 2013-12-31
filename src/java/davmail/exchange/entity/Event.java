package davmail.exchange.entity;

import davmail.BundleMessage;
import davmail.Settings;
import davmail.exception.DavMailException;
import davmail.exchange.ExchangeSession;
import davmail.exchange.VCalendar;
import davmail.io.MimeOutputStreamWriter;
import davmail.ui.NotificationDialog;
import org.apache.commons.httpclient.HttpException;
import org.apache.log4j.Logger;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimePart;
import java.io.*;
import java.util.Date;
import java.util.UUID;

/**
 * Calendar event object.
 */
public abstract class Event extends Item {
    private static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    protected final ExchangeSession exchangeSession;
    protected String contentClass;
    protected String subject;
    protected VCalendar vCalendar;

    private static int dumpIndex = 0;


    /**
     * @inheritDoc
     */
    public Event(ExchangeSession exchangeSession, String folderPath, String itemName, String contentClass, String itemBody, String etag, String noneMatch) throws IOException {
        super(folderPath, itemName, etag, noneMatch);
        this.exchangeSession = exchangeSession;
        this.contentClass = contentClass;
        fixICS(itemBody.getBytes("UTF-8"), false);
        // fix task item name
        if (vCalendar.isTodo() && this.getName().endsWith(".ics")) {
            this.itemName = itemName.substring(0, itemName.length() - 3) + "EML";
        }
    }

    /**
     * @inheritDoc
     */
    public Event(ExchangeSession exchangeSession) {
        this.exchangeSession = exchangeSession;
    }

    @Override
    public String getContentType() {
        return "text/calendar;charset=UTF-8";
    }

    @Override
    public String getBody() throws IOException {
        if (vCalendar == null) {
            fixICS(getEventContent(), true);
        }
        return vCalendar.toString();
    }

    protected HttpException buildHttpException(Exception e) {
        String message = "Unable to get event " + getName() + " subject: " + getSubject() + " at " + permanentUrl + ": " + e.getMessage();
        LOGGER.warn(message);
        return new HttpException(message);
    }

    /**
     * Retrieve item body from Exchange
     *
     * @return item content
     * @throws org.apache.commons.httpclient.HttpException on error
     */
    public abstract byte[] getEventContent() throws IOException;

    protected static final String TEXT_CALENDAR = "text/calendar";
    protected static final String APPLICATION_ICS = "application/ics";

    protected boolean isCalendarContentType(String contentType) {
        return TEXT_CALENDAR.regionMatches(true, 0, contentType, 0, TEXT_CALENDAR.length()) ||
                APPLICATION_ICS.regionMatches(true, 0, contentType, 0, APPLICATION_ICS.length());
    }

    protected MimePart getCalendarMimePart(MimeMultipart multiPart) throws IOException, MessagingException {
        MimePart bodyPart = null;
        for (int i = 0; i < multiPart.getCount(); i++) {
            String contentType = multiPart.getBodyPart(i).getContentType();
            if (isCalendarContentType(contentType)) {
                bodyPart = (MimePart) multiPart.getBodyPart(i);
                break;
            } else if (contentType.startsWith("multipart")) {
                Object content = multiPart.getBodyPart(i).getContent();
                if (content instanceof MimeMultipart) {
                    bodyPart = getCalendarMimePart((MimeMultipart) content);
                }
            }
        }

        return bodyPart;
    }

    /**
     * Load ICS content from MIME message input stream
     *
     * @param mimeInputStream mime message input stream
     * @return mime message ics attachment body
     * @throws java.io.IOException        on error
     * @throws javax.mail.MessagingException on error
     */
    protected byte[] getICS(InputStream mimeInputStream) throws IOException, MessagingException {
        byte[] result;
        MimeMessage mimeMessage = new MimeMessage(null, mimeInputStream);
        String[] contentClassHeader = mimeMessage.getHeader("Content-class");
        // task item, return null
        if (contentClassHeader != null && contentClassHeader.length > 0 && "urn:content-classes:task".equals(contentClassHeader[0])) {
            return null;
        }
        Object mimeBody = mimeMessage.getContent();
        MimePart bodyPart = null;
        if (mimeBody instanceof MimeMultipart) {
            bodyPart = getCalendarMimePart((MimeMultipart) mimeBody);
        } else if (isCalendarContentType(mimeMessage.getContentType())) {
            // no multipart, single body
            bodyPart = mimeMessage;
        }

        if (bodyPart != null) {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            bodyPart.getDataHandler().writeTo(baos);
            baos.close();
            result = baos.toByteArray();
        } else {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            mimeMessage.writeTo(baos);
            baos.close();
            throw new DavMailException("EXCEPTION_INVALID_MESSAGE_CONTENT", new String(baos.toByteArray(), "UTF-8"));
        }
        return result;
    }

    protected void fixICS(byte[] icsContent, boolean fromServer) throws IOException {
        if (LOGGER.isDebugEnabled() && fromServer) {
            dumpIndex++;
            String icsBody = new String(icsContent);
            dumpICS(icsBody, fromServer, false);
            LOGGER.debug("Vcalendar body received from server:\n" + icsBody);
        }
        vCalendar = new VCalendar(icsContent, exchangeSession.getEmail(), exchangeSession.getVTimezone());
        vCalendar.fixVCalendar(fromServer);
        if (LOGGER.isDebugEnabled() && !fromServer) {
            String resultString = vCalendar.toString();
            LOGGER.debug("Fixed Vcalendar body to server:\n" + resultString);
            dumpICS(resultString, fromServer, true);
        }
    }

    protected void dumpICS(String icsBody, boolean fromServer, boolean after) {
        String logFileDirectory = Settings.getLogFileDirectory();

        // additional setting to activate ICS dump (not available in GUI)
        int dumpMax = Settings.getIntProperty("davmail.dumpICS");
        if (dumpMax > 0) {
            if (dumpIndex > dumpMax) {
                // Delete the oldest dump file
                final int oldest = dumpIndex - dumpMax;
                try {
                    File[] oldestFiles = (new File(logFileDirectory)).listFiles(new FilenameFilter() {
                        public boolean accept(File dir, String name) {
                            if (name.endsWith(".ics")) {
                                int dashIndex = name.indexOf('-');
                                if (dashIndex > 0) {
                                    try {
                                        int fileIndex = Integer.parseInt(name.substring(0, dashIndex));
                                        return fileIndex < oldest;
                                    } catch (NumberFormatException nfe) {
                                        // ignore
                                    }
                                }
                            }
                            return false;
                        }
                    });
                    for (File file : oldestFiles) {
                        if (!file.delete()) {
                            LOGGER.warn("Unable to delete " + file.getAbsolutePath());
                        }
                    }
                } catch (Exception ex) {
                    LOGGER.warn("Error deleting ics dump: " + ex.getMessage());
                }
            }

            StringBuilder filePath = new StringBuilder();
            filePath.append(logFileDirectory).append('/')
                    .append(dumpIndex)
                    .append(after ? "-to" : "-from")
                    .append((after ^ fromServer) ? "-server" : "-client")
                    .append(".ics");
            if ((icsBody != null) && (icsBody.length() > 0)) {
                FileWriter fileWriter = null;
                try {
                    fileWriter = new FileWriter(filePath.toString());
                    fileWriter.write(icsBody);
                } catch (IOException e) {
                    LOGGER.error(e);
                } finally {
                    if (fileWriter != null) {
                        try {
                            fileWriter.close();
                        } catch (IOException e) {
                            LOGGER.error(e);
                        }
                    }
                }


            }
        }

    }

    /**
     * Build Mime body for event or event message.
     *
     * @return mimeContent as byte array or null
     * @throws java.io.IOException on error
     */
    public byte[] createMimeContent() throws IOException {
        String boundary = UUID.randomUUID().toString();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        MimeOutputStreamWriter writer = new MimeOutputStreamWriter(baos);

        writer.writeHeader("Content-Transfer-Encoding", "7bit");
        writer.writeHeader("Content-class", contentClass);
        // append date
        writer.writeHeader("Date", new Date());

        // Make sure invites have a proper subject line
        String vEventSubject = vCalendar.getFirstVeventPropertyValue("SUMMARY");
        if (vEventSubject == null) {
            vEventSubject = BundleMessage.format("MEETING_REQUEST");
        }

        // Write a part of the message that contains the
        // ICS description so that invites contain the description text
        String description = vCalendar.getFirstVeventPropertyValue("DESCRIPTION");

        // handle notifications
        if ("urn:content-classes:calendarmessage".equals(contentClass)) {
            // need to parse attendees and organizer to build recipients
            VCalendar.Recipients recipients = vCalendar.getRecipients(true);
            String to;
            String cc;
            String notificationSubject;
            if (exchangeSession.email.equalsIgnoreCase(recipients.organizer)) {
                // current user is organizer => notify all
                to = recipients.attendees;
                cc = recipients.optionalAttendees;
                notificationSubject = getSubject();
            } else {
                String status = vCalendar.getAttendeeStatus();
                // notify only organizer
                to = recipients.organizer;
                cc = null;
                notificationSubject = (status != null) ? (BundleMessage.format(status) + vEventSubject) : getSubject();
                description = "";
            }

            // Allow end user notification edit
            if (Settings.getBooleanProperty("davmail.caldavEditNotifications")) {
                // create notification edit dialog
                NotificationDialog notificationDialog = new NotificationDialog(to,
                        cc, notificationSubject, description);
                if (!notificationDialog.getSendNotification()) {
                    LOGGER.debug("Notification canceled by user");
                    return null;
                }
                // get description from dialog
                to = notificationDialog.getTo();
                cc = notificationDialog.getCc();
                notificationSubject = notificationDialog.getSubject();
                description = notificationDialog.getBody();
            }

            // do not send notification if no recipients found
            if ((to == null || to.length() == 0) && (cc == null || cc.length() == 0)) {
                return null;
            }

            writer.writeHeader("To", to);
            writer.writeHeader("Cc", cc);
            writer.writeHeader("Subject", notificationSubject);


            if (LOGGER.isDebugEnabled()) {
                StringBuilder logBuffer = new StringBuilder("Sending notification ");
                if (to != null) {
                    logBuffer.append("to: ").append(to);
                }
                if (cc != null) {
                    logBuffer.append("cc: ").append(cc);
                }
                LOGGER.debug(logBuffer.toString());
            }
        } else {
            // need to parse attendees and organizer to build recipients
            VCalendar.Recipients recipients = vCalendar.getRecipients(false);
            // storing appointment, full recipients header
            if (recipients.attendees != null) {
                writer.writeHeader("To", recipients.attendees);
            } else {
                // use current user as attendee
                writer.writeHeader("To", exchangeSession.email);
            }
            writer.writeHeader("Cc", recipients.optionalAttendees);

            if (recipients.organizer != null) {
                writer.writeHeader("From", recipients.organizer);
            } else {
                writer.writeHeader("From", exchangeSession.email);
            }
        }
        if (vCalendar.getMethod() == null) {
            vCalendar.setPropertyValue("METHOD", "REQUEST");
        }
        writer.writeHeader("MIME-Version", "1.0");
        writer.writeHeader("Content-Type", "multipart/alternative;\r\n" +
                "\tboundary=\"----=_NextPart_" + boundary + '\"');
        writer.writeLn();
        writer.writeLn("This is a multi-part message in MIME format.");
        writer.writeLn();
        writer.writeLn("------=_NextPart_" + boundary);

        if (description != null && description.length() > 0) {
            writer.writeHeader("Content-Type", "text/plain;\r\n" +
                    "\tcharset=\"utf-8\"");
            writer.writeHeader("content-transfer-encoding", "8bit");
            writer.writeLn();
            writer.flush();
            baos.write(description.getBytes("UTF-8"));
            writer.writeLn();
            writer.writeLn("------=_NextPart_" + boundary);
        }
        writer.writeHeader("Content-class", contentClass);
        writer.writeHeader("Content-Type", "text/calendar;\r\n" +
                "\tmethod=" + vCalendar.getMethod() + ";\r\n" +
                "\tcharset=\"utf-8\""
        );
        writer.writeHeader("Content-Transfer-Encoding", "8bit");
        writer.writeLn();
        writer.flush();
        baos.write(vCalendar.toString().getBytes("UTF-8"));
        writer.writeLn();
        writer.writeLn("------=_NextPart_" + boundary + "--");
        writer.close();
        return baos.toByteArray();
    }

    /**
     * Create or update item
     *
     * @return action result
     * @throws java.io.IOException on error
     */
    public abstract ItemResult createOrUpdate() throws IOException;

    public String getSubject() {
        return subject;
    }
}
