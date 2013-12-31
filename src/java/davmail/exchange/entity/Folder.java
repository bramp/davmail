package davmail.exchange.entity;

import davmail.exchange.condition.Condition;
import davmail.exchange.ExchangeSession;
import davmail.exchange.MessageList;
import org.apache.log4j.Logger;

import java.io.IOException;
import java.util.Collections;
import java.util.HashMap;
import java.util.TreeMap;

/**
 * Exchange folder with IMAP properties
 */
public class Folder {

    private static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    /**
     * Logical (IMAP) folder path.
     */
    public String folderPath;

    /**
     * Display Name.
     */
    public String displayName;
    /**
     * Folder class (PR_CONTAINER_CLASS).
     */
    public String folderClass;
    /**
     * Folder message count.
     */
    public int count;
    /**
     * Folder unread message count.
     */
    public int unreadCount;
    /**
     * true if folder has subfolders (DAV:hassubs).
     */
    public boolean hasChildren;
    /**
     * true if folder has no subfolders (DAV:nosubs).
     */
    public boolean noInferiors;
    /**
     * Folder content tag (to detect folder content changes).
     */
    public String ctag;
    /**
     * Folder etag (to detect folder object changes).
     */
    public String etag;
    /**
     * Next IMAP uid
     */
    public long uidNext;
    /**
     * recent count
     */
    public int recent;

    public final ExchangeSession exchangeSession;

    /**
     * Folder message list, empty before loadMessages call.
     */
    public MessageList messages;
    /**
     * Permanent uid (PR_SEARCH_KEY) to IMAP UID map.
     */
    private final HashMap<String, Long> permanentUrlToImapUidMap = new HashMap<String, Long>();

    public Folder(ExchangeSession exchangeSession) {
        this.exchangeSession = exchangeSession;
    }

    /**
     * Get IMAP folder flags.
     *
     * @return folder flags in IMAP format
     */
    public String getFlags() {
        if (noInferiors) {
            return "\\NoInferiors";
        } else if (hasChildren) {
            return "\\HasChildren";
        } else {
            return "\\HasNoChildren";
        }
    }

    /**
     * Load folder messages.
     *
     * @throws java.io.IOException on error
     */
    public void loadMessages() throws IOException {
        messages = exchangeSession.searchMessages(folderPath, null);
        fixUids(messages);
        recent = 0;
        for (Message message : messages) {
            if (message.recent) {
                recent++;
            }
        }
        long computedUidNext = 1;
        if (!messages.isEmpty()) {
            computedUidNext = messages.get(messages.size() - 1).getImapUid() + 1;
        }
        if (computedUidNext > uidNext) {
            uidNext = computedUidNext;
        }
    }

    /**
     * Search messages in folder matching query.
     *
     * @param condition search query
     * @return message list
     * @throws java.io.IOException on error
     */
    public MessageList searchMessages(Condition condition) throws IOException {
        MessageList localMessages = exchangeSession.searchMessages(folderPath, condition);
        fixUids(localMessages);
        return localMessages;
    }

    /**
     * Restore previous uids changed by a PROPPATCH (flag change).
     *
     * @param messages message list
     */
    protected void fixUids(MessageList messages) {
        boolean sortNeeded = false;
        for (Message message : messages) {
            if (permanentUrlToImapUidMap.containsKey(message.getPermanentId())) {
                long previousUid = permanentUrlToImapUidMap.get(message.getPermanentId());
                if (message.getImapUid() != previousUid) {
                    LOGGER.debug("Restoring IMAP uid " + message.getImapUid() + " -> " + previousUid + " for message " + message.getPermanentId());
                    message.setImapUid(previousUid);
                    sortNeeded = true;
                }
            } else {
                // add message to uid map
                permanentUrlToImapUidMap.put(message.getPermanentId(), message.getImapUid());
            }
        }
        if (sortNeeded) {
            Collections.sort(messages);
        }
    }

    /**
     * Folder message count.
     *
     * @return message count
     */
    public int count() {
        if (messages == null) {
            return count;
        } else {
            return messages.size();
        }
    }

    /**
     * Compute IMAP uidnext.
     *
     * @return max(messageuids)+1
     */
    public long getUidNext() {
        return uidNext;
    }

    /**
     * Get message at index.
     *
     * @param index message index
     * @return message
     */
    public Message get(int index) {
        return messages.get(index);
    }

    /**
     * Get current folder messages imap uids and flags
     *
     * @return imap uid list
     */
    public TreeMap<Long, String> getImapFlagMap() {
        TreeMap<Long, String> imapFlagMap = new TreeMap<Long, String>();
        for (Message message : messages) {
            imapFlagMap.put(message.getImapUid(), message.getImapFlags());
        }
        return imapFlagMap;
    }

    /**
     * Calendar folder flag.
     *
     * @return true if this is a calendar folder
     */
    public boolean isCalendar() {
        return "IPF.Appointment".equals(folderClass);
    }

    /**
     * Contact folder flag.
     *
     * @return true if this is a calendar folder
     */
    public boolean isContact() {
        return "IPF.Contact".equals(folderClass);
    }

    /**
     * Task folder flag.
     *
     * @return true if this is a task folder
     */
    public boolean isTask() {
        return "IPF.Task".equals(folderClass);
    }

    /**
     * drop cached message
     */
    public void clearCache() {
        messages.cachedMimeBody = null;
        messages.cachedMimeMessage = null;
        messages.cachedMessageImapUid = 0;
    }
}
