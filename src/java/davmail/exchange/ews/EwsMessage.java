package davmail.exchange.ews;

import davmail.exchange.ExchangeSession;
import org.apache.log4j.Logger;

import java.io.ByteArrayInputStream;
import java.io.InputStream;

class EwsMessage extends davmail.exchange.entity.Message {
    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private final EwsExchangeSession ewsExchangeSession;

    // message item id
    ItemId itemId;

    public EwsMessage(EwsExchangeSession ewsExchangeSession) {
        super(ewsExchangeSession);
        this.ewsExchangeSession = ewsExchangeSession;
    }

    @Override
    public String getPermanentId() {
        return itemId.id;
    }

    @Override
    protected InputStream getMimeHeaders() {
        InputStream result = null;
        try {
            GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, itemId, false);
            getItemMethod.addAdditionalProperty(Field.get("messageheaders"));
            getItemMethod.addAdditionalProperty(Field.get("from"));
            ewsExchangeSession.executeMethod(getItemMethod);
            EWSMethod.Item item = getItemMethod.getResponseItem();

            String messageHeaders = item.get(Field.get("messageheaders").getResponseName());
            if (messageHeaders != null
                    // workaround for broken message headers on Exchange 2010
                    && messageHeaders.toLowerCase().contains("message-id:")) {
                // workaround for messages in Sent folder
                if (messageHeaders.indexOf("From:") < 0) {
                    String from = item.get(Field.get("from").getResponseName());
                    messageHeaders = "From: " + from + "\n" + messageHeaders;
                }

                result = new ByteArrayInputStream(messageHeaders.getBytes("UTF-8"));
            }
        } catch (Exception e) {
            LOGGER.warn(e.getMessage());
        }

        return result;
    }
}
