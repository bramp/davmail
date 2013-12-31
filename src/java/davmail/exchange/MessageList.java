package davmail.exchange;

import davmail.exchange.entity.Message;

import javax.mail.internet.MimeMessage;
import javax.mail.util.SharedByteArrayInputStream;
import java.util.ArrayList;

/**
 * Message list, includes a single messsage cache
 */
public class MessageList extends ArrayList<Message> {


    private static final long serialVersionUID = 0;

    /**
     * Cached message content parsed in a MIME message.
     */
    public transient MimeMessage cachedMimeMessage;
    /**
     * Cached message uid.
     */
    public transient long cachedMessageImapUid;
    /**
     * Cached unparsed message
     */
    public transient SharedByteArrayInputStream cachedMimeBody;

}
