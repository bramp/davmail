package davmail.exchange.dav;

import davmail.exchange.ExchangeSession;
import org.apache.log4j.Logger;

import java.io.ByteArrayInputStream;
import java.io.InputStream;

/**
 * @inheritDoc
 */
public class DavMessage extends davmail.exchange.entity.Message {

    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private final DavExchangeSession davExchangeSession;

    public DavMessage(DavExchangeSession davExchangeSession) {
        super(davExchangeSession);
        this.davExchangeSession = davExchangeSession;
    }

    @Override
    public String getPermanentId() {
        return permanentUrl;
    }

    @Override
    protected InputStream getMimeHeaders() {
        InputStream input = null;
        try {
            String messageHeaders = davExchangeSession.getItemProperty(permanentUrl, "messageheaders");
            if (messageHeaders != null) {
                final String MS_HEADER = "Microsoft Mail Internet Headers Version 2.0";
                if (messageHeaders.startsWith(MS_HEADER)) {
                    messageHeaders = messageHeaders.substring(MS_HEADER.length());
                    if (!messageHeaders.isEmpty() && messageHeaders.charAt(0) == '\r') {
                        messageHeaders = messageHeaders.substring(1);
                    }
                    if (!messageHeaders.isEmpty() && messageHeaders.charAt(0) == '\n') {
                        messageHeaders = messageHeaders.substring(1);
                    }
                }
                // workaround for messages in Sent folder
                if (messageHeaders.indexOf("From:") < 0) {
                    String from = davExchangeSession.getItemProperty(permanentUrl, "from");
                    messageHeaders = "From: "+from+"\n"+messageHeaders;
                }
                input = new ByteArrayInputStream(messageHeaders.getBytes("UTF-8"));
            }
        } catch (Exception e) {
            LOGGER.warn(e.getMessage());
        }
        return input;
    }
}
