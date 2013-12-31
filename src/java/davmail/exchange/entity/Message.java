package davmail.exchange.entity;

import davmail.exchange.ExchangeSession;
import davmail.exchange.MessageList;
import davmail.util.StringUtil;
import org.apache.log4j.Logger;

import javax.mail.MessagingException;
import javax.mail.internet.InternetHeaders;
import javax.mail.internet.MimeMessage;
import javax.mail.util.SharedByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Set;

/**
 * Exchange message.
 */
public abstract class Message implements Comparable<Message> {

    private static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private ExchangeSession exchangeSession;
    /**
     * enclosing message list
     */
    public MessageList messageList;
    /**
     * Message url.
     */
    public String messageUrl;
    /**
     * Message permanent url (does not change on message move).
     */
    public String permanentUrl;
    /**
     * Message uid.
     */
    public String uid;
    /**
     * Message content class.
     */
    public String contentClass;
    /**
     * Message keywords (categories).
     */
    public String keywords;
    /**
     * Message IMAP uid, unique in folder (x0e230003).
     */
    public long imapUid;
    /**
     * MAPI message size.
     */
    public int size;
    /**
     * Message date (urn:schemas:mailheader:date).
     */
    public String date;

    /**
     * Message flag: read.
     */
    public boolean read;
    /**
     * Message flag: deleted.
     */
    public boolean deleted;
    /**
     * Message flag: junk.
     */
    public boolean junk;
    /**
     * Message flag: flagged.
     */
    public boolean flagged;
    /**
     * Message flag: recent.
     */
    public boolean recent;
    /**
     * Message flag: draft.
     */
    public boolean draft;
    /**
     * Message flag: answered.
     */
    public boolean answered;
    /**
     * Message flag: fowarded.
     */
    public boolean forwarded;

    /**
     * Unparsed message content.
     */
    protected SharedByteArrayInputStream mimeBody;

    /**
     * Message content parsed in a MIME message.
     */
    protected MimeMessage mimeMessage;

    public Message(ExchangeSession exchangeSession) {
        this.exchangeSession = exchangeSession;
    }

    /**
     * Get permanent message id.
     * permanentUrl over WebDav or IitemId over EWS
     *
     * @return permanent id
     */
    public abstract String getPermanentId();

    /**
     * IMAP uid , unique in folder (x0e230003)
     *
     * @return IMAP uid
     */
    public long getImapUid() {
        return imapUid;
    }

    /**
     * Set IMAP uid.
     *
     * @param imapUid new uid
     */
    public void setImapUid(long imapUid) {
        this.imapUid = imapUid;
    }

    /**
     * Exchange uid.
     *
     * @return uid
     */
    public String getUid() {
        return uid;
    }

    /**
     * Return message flags in IMAP format.
     *
     * @return IMAP flags
     */
    public String getImapFlags() {
        StringBuilder buffer = new StringBuilder();
        if (read) {
            buffer.append("\\Seen ");
        }
        if (deleted) {
            buffer.append("\\Deleted ");
        }
        if (recent) {
            buffer.append("\\Recent ");
        }
        if (flagged) {
            buffer.append("\\Flagged ");
        }
        if (junk) {
            buffer.append("Junk ");
        }
        if (draft) {
            buffer.append("\\Draft ");
        }
        if (answered) {
            buffer.append("\\Answered ");
        }
        if (forwarded) {
            buffer.append("$Forwarded ");
        }
        if (keywords != null) {
            for (String keyword : keywords.split(",")) {
                buffer.append(exchangeSession.convertKeywordToFlag(keyword)).append(" ");
            }
        }
        return buffer.toString().trim();
    }

    /**
     * Load message content in a Mime message
     *
     * @throws java.io.IOException        on error
     * @throws javax.mail.MessagingException on error
     */
    public void loadMimeMessage() throws IOException, MessagingException {
        if (mimeMessage == null) {
            // try to get message content from cache
            if (this.imapUid == messageList.cachedMessageImapUid) {
                mimeBody = messageList.cachedMimeBody;
                mimeMessage = messageList.cachedMimeMessage;
                LOGGER.debug("Got message content for " + imapUid + " from cache");
            } else {
                // load and parse message
                mimeBody = new SharedByteArrayInputStream(exchangeSession.getContent(this));
                mimeMessage = new MimeMessage(null, mimeBody);
                mimeBody.reset();
                // workaround for Exchange 2003 ActiveSync bug
                if (mimeMessage.getHeader("MAIL FROM") != null) {
                    mimeBody = (SharedByteArrayInputStream) mimeMessage.getRawInputStream();
                    mimeMessage = new MimeMessage(null, mimeBody);
                    mimeBody.reset();
                }
                LOGGER.debug("Downloaded full message content for IMAP UID " + imapUid + " (" + mimeBody.available() + " bytes)");
            }
        }
    }

    /**
     * Get message content as a Mime message.
     *
     * @return mime message
     * @throws java.io.IOException        on error
     * @throws javax.mail.MessagingException on error
     */
    public MimeMessage getMimeMessage() throws IOException, MessagingException {
        loadMimeMessage();
        return mimeMessage;
    }

    public Enumeration getMatchingHeaderLinesFromHeaders(String[] headerNames) throws MessagingException, IOException {
        Enumeration result = null;
        if (mimeMessage == null) {
            // message not loaded, try to get headers only
            InputStream headers = getMimeHeaders();
            if (headers != null) {
                InternetHeaders internetHeaders = new InternetHeaders(headers);
                if (internetHeaders.getHeader("Subject") == null) {
                    // invalid header content
                    return null;
                }
                if (headerNames == null) {
                    result = internetHeaders.getAllHeaderLines();
                } else {
                    result = internetHeaders.getMatchingHeaderLines(headerNames);
                }
            }
        }
        return result;
    }

    public Enumeration getMatchingHeaderLines(String[] headerNames) throws MessagingException, IOException {
        Enumeration result = getMatchingHeaderLinesFromHeaders(headerNames);
        if (result == null) {
            if (headerNames == null) {
                result = getMimeMessage().getAllHeaderLines();
            } else {
                result = getMimeMessage().getMatchingHeaderLines(headerNames);
            }

        }
        return result;
    }

    protected abstract InputStream getMimeHeaders();

    /**
     * Get message body size.
     *
     * @return mime message size
     * @throws java.io.IOException        on error
     * @throws javax.mail.MessagingException on error
     */
    public int getMimeMessageSize() throws IOException, MessagingException {
        loadMimeMessage();
        mimeBody.reset();
        return mimeBody.available();
    }

    /**
     * Get message body input stream.
     *
     * @return mime message InputStream
     * @throws java.io.IOException        on error
     * @throws javax.mail.MessagingException on error
     */
    public InputStream getRawInputStream() throws IOException, MessagingException {
        loadMimeMessage();
        mimeBody.reset();
        return mimeBody;
    }


    /**
     * Drop mime message to avoid keeping message content in memory,
     * keep a single message in MessageList cache to handle chunked fetch.
     */
    public void dropMimeMessage() {
        // update single message cache
        if (mimeMessage != null) {
            messageList.cachedMessageImapUid = imapUid;
            messageList.cachedMimeBody = mimeBody;
            messageList.cachedMimeMessage = mimeMessage;
        }
        // drop curent message body to save memory
        mimeMessage = null;
        mimeBody = null;
    }

    public boolean isLoaded() {
        return mimeMessage != null;
    }

    /**
     * Delete message.
     *
     * @throws java.io.IOException on error
     */
    public void delete() throws IOException {
        exchangeSession.deleteMessage(this);
    }

    /**
     * Move message to trash, mark message read.
     *
     * @throws java.io.IOException on error
     */
    public void moveToTrash() throws IOException {
        markRead();

        exchangeSession.moveToTrash(this);
    }

    /**
     * Mark message as read.
     *
     * @throws java.io.IOException on error
     */
    public void markRead() throws IOException {
        HashMap<String, String> properties = new HashMap<String, String>();
        properties.put("read", "1");
        exchangeSession.updateMessage(this, properties);
    }

    /**
     * Comparator to sort messages by IMAP uid
     *
     * @param message other message
     * @return imapUid comparison result
     */
    public int compareTo(Message message) {
        long compareValue = (imapUid - message.imapUid);
        if (compareValue > 0) {
            return 1;
        } else if (compareValue < 0) {
            return -1;
        } else {
            return 0;
        }
    }

    /**
     * Override equals, compare IMAP uids
     *
     * @param message other message
     * @return true if IMAP uids are equal
     */
    @Override
    public boolean equals(Object message) {
        return message instanceof Message && imapUid == ((Message) message).imapUid;
    }

    /**
     * Override hashCode, return imapUid hashcode.
     *
     * @return imapUid hashcode
     */
    @Override
    public int hashCode() {
        return (int) (imapUid ^ (imapUid >>> 32));
    }

    public String removeFlag(String flag) {
        if (keywords != null) {
            final String exchangeFlag = exchangeSession.convertFlagToKeyword(flag);
            Set<String> keywordSet = new HashSet<String>();
            String[] keywordArray = keywords.split(",");
            for (String value : keywordArray) {
                if (!value.equalsIgnoreCase(exchangeFlag)) {
                    keywordSet.add(value);
                }
            }
            keywords = StringUtil.join(keywordSet, ",");
        }
        return keywords;
    }

    public String addFlag(String flag) {
        final String exchangeFlag = exchangeSession.convertFlagToKeyword(flag);
        HashSet<String> keywordSet = new HashSet<String>();
        boolean hasFlag = false;
        if (keywords != null) {
            String[] keywordArray = keywords.split(",");
            for (String value : keywordArray) {
                keywordSet.add(value);
                if (value.equalsIgnoreCase(exchangeFlag)) {
                    hasFlag = true;
                }
            }
        }
        if (!hasFlag) {
            keywordSet.add(exchangeFlag);
        }
        keywords = StringUtil.join(keywordSet, ",");
        return keywords;
    }

    public String setFlags(HashSet<String> flags) {
        HashSet<String> keywordSet = new HashSet<String>();
        for (String flag : flags) {
            keywordSet.add(exchangeSession.convertFlagToKeyword(flag));
        }
        keywords = StringUtil.join(keywordSet, ",");
        return keywords;
    }

}
