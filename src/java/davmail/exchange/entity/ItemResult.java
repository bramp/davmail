package davmail.exchange.entity;

/**
 * Event result object to hold HTTP status and event etag from an event creation/update.
 */
public class ItemResult {
    /**
     * HTTP status
     */
    public int status;
    /**
     * Event etag from response HTTP header
     */
    public String etag;
}
