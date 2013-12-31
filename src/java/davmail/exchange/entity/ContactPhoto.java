package davmail.exchange.entity;

/**
 * Contact picture
 */
public class ContactPhoto {
    /**
     * Contact picture content type (always image/jpeg on read)
     */
    public String contentType;
    /**
     * Base64 encoded picture content
     */
    public String content;
}
