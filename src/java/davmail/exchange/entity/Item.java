package davmail.exchange.entity;

import java.io.IOException;
import java.util.HashMap;

/**
 * Generic folder item.
 */
public abstract class Item extends HashMap<String, String> {
    protected String folderPath;
    protected String itemName;
    protected String permanentUrl;
    /**
     * Display name.
     */
    public String displayName;
    /**
     * item etag
     */
    public String etag;
    protected String noneMatch;

    /**
     * Build item instance.
     *
     * @param folderPath folder path
     * @param itemName   item name class
     * @param etag       item etag
     * @param noneMatch  none match flag
     */
    public Item(String folderPath, String itemName, String etag, String noneMatch) {
        this.folderPath = folderPath;
        this.itemName = itemName;
        this.etag = etag;
        this.noneMatch = noneMatch;
    }

    /**
     * Default constructor.
     */
    Item() {
    }

    /**
     * Return item content type
     *
     * @return content type
     */
    public abstract String getContentType();

    /**
     * Retrieve item body from Exchange
     *
     * @return item body
     * @throws org.apache.commons.httpclient.HttpException on error
     */
    public abstract String getBody() throws IOException;

    /**
     * Get event name (file name part in URL).
     *
     * @return event name
     */
    public String getName() {
        return itemName;
    }

    /**
     * Get event etag (last change tag).
     *
     * @return event etag
     */
    public String getEtag() {
        return etag;
    }

    /**
     * Set item href.
     *
     * @param href item href
     */
    public void setHref(String href) {
        int index = href.lastIndexOf('/');
        if (index >= 0) {
            folderPath = href.substring(0, index);
            itemName = href.substring(index + 1);
        } else {
            throw new IllegalArgumentException(href);
        }
    }

    /**
     * Return item href.
     *
     * @return item href
     */
    public String getHref() {
        return folderPath + '/' + itemName;
    }
}
