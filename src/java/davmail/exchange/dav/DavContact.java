package davmail.exchange.dav;

import davmail.exception.DavMailException;
import davmail.exchange.ExchangeSession;
import davmail.exchange.entity.ItemResult;
import davmail.util.IOUtil;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.httpclient.HttpStatus;
import org.apache.commons.httpclient.URIException;
import org.apache.commons.httpclient.methods.ByteArrayRequestEntity;
import org.apache.commons.httpclient.methods.DeleteMethod;
import org.apache.commons.httpclient.methods.HeadMethod;
import org.apache.commons.httpclient.methods.PutMethod;
import org.apache.commons.httpclient.util.URIUtil;
import org.apache.jackrabbit.webdav.MultiStatusResponse;
import org.apache.jackrabbit.webdav.property.DavPropertySet;
import org.apache.log4j.Logger;

import java.io.IOException;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * @inheritDoc
 */
public class DavContact extends davmail.exchange.entity.Contact {
    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private final DavExchangeSession davExchangeSession;

    /**
     * Build Contact instance from multistatusResponse info
     *
     * @param multiStatusResponse response
     * @throws org.apache.commons.httpclient.URIException     on error
     * @throws davmail.exception.DavMailException on error
     */
    public DavContact(DavExchangeSession davExchangeSession, MultiStatusResponse multiStatusResponse) throws URIException, DavMailException {
        super(davExchangeSession);

        this.davExchangeSession = davExchangeSession;
        setHref(URIUtil.decode(multiStatusResponse.getHref()));
        DavPropertySet properties = multiStatusResponse.getProperties(HttpStatus.SC_OK);
        permanentUrl = davExchangeSession.getURLPropertyIfExists(properties, "permanenturl");
        etag = davExchangeSession.getPropertyIfExists(properties, "etag");
        displayName = davExchangeSession.getPropertyIfExists(properties, "displayname");
        for (String attributeName : ExchangeSession.CONTACT_ATTRIBUTES) {
            String value = davExchangeSession.getPropertyIfExists(properties, attributeName);
            if (value != null) {
                if ("bday".equals(attributeName) || "anniversary".equals(attributeName)
                        || "lastmodified".equals(attributeName) || "datereceived".equals(attributeName)) {
                    value = davExchangeSession.convertDateFromExchange(value);
                } else if ("haspicture".equals(attributeName) || "private".equals(attributeName)) {
                    value = "1".equals(value) ? "true" : "false";
                }
                put(attributeName, value);
            }
        }
    }

    /**
     * @inheritDoc
     */
    public DavContact(DavExchangeSession davExchangeSession, String folderPath, String itemName, Map<String, String> properties, String etag, String noneMatch) {
        super(davExchangeSession, folderPath, itemName, properties, etag, noneMatch);
        this.davExchangeSession = davExchangeSession;
    }

    /**
     * Default constructor for galFind
     */
    public DavContact(DavExchangeSession davExchangeSession) {
        super(davExchangeSession);
        this.davExchangeSession = davExchangeSession;
    }

    protected Set<PropertyValue> buildProperties() {
        Set<PropertyValue> propertyValues = new HashSet<PropertyValue>();
        for (Map.Entry<String, String> entry : entrySet()) {
            String key = entry.getKey();
            if (!"photo".equals(key)) {
                propertyValues.add(Field.createPropertyValue(key, entry.getValue()));
                if (key.startsWith("email")) {
                    propertyValues.add(Field.createPropertyValue(key + "type", "SMTP"));
                }
            }
        }

        return propertyValues;
    }

    protected ExchangePropPatchMethod internalCreateOrUpdate(String encodedHref) throws IOException {
        ExchangePropPatchMethod propPatchMethod = new ExchangePropPatchMethod(encodedHref, buildProperties());
        propPatchMethod.setRequestHeader("Translate", "f");
        if (etag != null) {
            propPatchMethod.setRequestHeader("If-Match", etag);
        }
        if (noneMatch != null) {
            propPatchMethod.setRequestHeader("If-None-Match", noneMatch);
        }
        try {
            exchangeSession.getHttpClient().executeMethod(propPatchMethod);
        } finally {
            propPatchMethod.releaseConnection();
        }
        return propPatchMethod;
    }

    /**
     * Create or update contact
     *
     * @return action result
     * @throws java.io.IOException on error
     */
    public ItemResult createOrUpdate() throws IOException {
        String encodedHref = URIUtil.encodePath(getHref());
        ExchangePropPatchMethod propPatchMethod = internalCreateOrUpdate(encodedHref);
        int status = propPatchMethod.getStatusCode();
        if (status == HttpStatus.SC_MULTI_STATUS) {
            status = propPatchMethod.getResponseStatusCode();
            //noinspection VariableNotUsedInsideIf
            if (status == HttpStatus.SC_CREATED) {
                LOGGER.debug("Created contact " + encodedHref);
            } else {
                LOGGER.debug("Updated contact " + encodedHref);
            }
        } else if (status == HttpStatus.SC_NOT_FOUND) {
            LOGGER.debug("Contact not found at " + encodedHref + ", searching permanenturl by urlcompname");
            // failover, search item by urlcompname
            MultiStatusResponse[] responses = davExchangeSession.searchItems(folderPath, DavExchangeSession.EVENT_REQUEST_PROPERTIES, exchangeSession.isEqualTo("urlcompname", exchangeSession.convertItemNameToEML(itemName)), DavExchangeSession.FolderQueryTraversal.Shallow, 1);
            if (responses.length == 1) {
                encodedHref = davExchangeSession.getPropertyIfExists(responses[0].getProperties(HttpStatus.SC_OK), "permanenturl");
                LOGGER.warn("Contact found, permanenturl is " + encodedHref);
                propPatchMethod = internalCreateOrUpdate(encodedHref);
                status = propPatchMethod.getStatusCode();
                if (status == HttpStatus.SC_MULTI_STATUS) {
                    status = propPatchMethod.getResponseStatusCode();
                    LOGGER.debug("Updated contact " + encodedHref);
                } else {
                    LOGGER.warn("Unable to create or update contact " + status + ' ' + propPatchMethod.getStatusLine());
                }
            }

        } else {
            LOGGER.warn("Unable to create or update contact " + status + ' ' + propPatchMethod.getStatusLine());
        }
        ItemResult itemResult = new ItemResult();
        // 440 means forbidden on Exchange
        if (status == 440) {
            status = HttpStatus.SC_FORBIDDEN;
        }
        itemResult.status = status;

        if (status == HttpStatus.SC_OK || status == HttpStatus.SC_CREATED) {
            String contactPictureUrl = URIUtil.encodePath(getHref() + "/ContactPicture.jpg");
            String photo = get("photo");
            if (photo != null) {
                // need to update photo
                byte[] resizedImageBytes = IOUtil.resizeImage(Base64.decodeBase64(photo.getBytes()), 90);

                final PutMethod putmethod = new PutMethod(contactPictureUrl);
                putmethod.setRequestHeader("Overwrite", "t");
                putmethod.setRequestHeader("Content-Type", "image/jpeg");
                putmethod.setRequestEntity(new ByteArrayRequestEntity(resizedImageBytes, "image/jpeg"));
                try {
                    status = exchangeSession.getHttpClient().executeMethod(putmethod);
                    if (status != HttpStatus.SC_OK && status != HttpStatus.SC_CREATED) {
                        throw new IOException("Unable to update contact picture: " + status + ' ' + putmethod.getStatusLine());
                    }
                } catch (IOException e) {
                    LOGGER.error("Error in contact photo create or update", e);
                    throw e;
                } finally {
                    putmethod.releaseConnection();
                }

                Set<PropertyValue> picturePropertyValues = new HashSet<PropertyValue>();
                picturePropertyValues.add(Field.createPropertyValue("attachmentContactPhoto", "true"));
                // picturePropertyValues.add(Field.createPropertyValue("renderingPosition", "-1"));
                picturePropertyValues.add(Field.createPropertyValue("attachExtension", ".jpg"));

                final ExchangePropPatchMethod attachmentPropPatchMethod = new ExchangePropPatchMethod(contactPictureUrl, picturePropertyValues);
                try {
                    status = exchangeSession.getHttpClient().executeMethod(attachmentPropPatchMethod);
                    if (status != HttpStatus.SC_MULTI_STATUS) {
                        LOGGER.error("Error in contact photo create or update: " + attachmentPropPatchMethod.getStatusCode());
                        throw new IOException("Unable to update contact picture");
                    }
                } finally {
                    attachmentPropPatchMethod.releaseConnection();
                }

            } else {
                // try to delete picture
                DeleteMethod deleteMethod = new DeleteMethod(contactPictureUrl);
                try {
                    status = exchangeSession.getHttpClient().executeMethod(deleteMethod);
                    if (status != HttpStatus.SC_OK && status != HttpStatus.SC_NOT_FOUND) {
                        LOGGER.error("Error in contact photo delete: " + status);
                        throw new IOException("Unable to delete contact picture");
                    }
                } finally {
                    deleteMethod.releaseConnection();
                }
            }
            // need to retrieve new etag
            HeadMethod headMethod = new HeadMethod(URIUtil.encodePath(getHref()));
            try {
                exchangeSession.getHttpClient().executeMethod(headMethod);
                if (headMethod.getResponseHeader("ETag") != null) {
                    itemResult.etag = headMethod.getResponseHeader("ETag").getValue();
                }
            } finally {
                headMethod.releaseConnection();
            }
        }
        return itemResult;

    }

}
