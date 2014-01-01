package davmail.exchange.ews;

import davmail.exception.DavMailException;
import davmail.exchange.ExchangeSession;
import davmail.exchange.ExchangeVersion;
import davmail.exchange.entity.ItemResult;
import davmail.util.IOUtil;
import davmail.util.StringUtil;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.httpclient.HttpStatus;
import org.apache.log4j.Logger;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

class EwsContact extends davmail.exchange.entity.Contact {
    protected static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    private final EwsExchangeSession ewsExchangeSession;

    // item id
    ItemId itemId;

    EwsContact(EwsExchangeSession ewsExchangeSession, EWSMethod.Item response) throws DavMailException {
        super(ewsExchangeSession);

        this.ewsExchangeSession = ewsExchangeSession;
        itemId = new ItemId(response);

        permanentUrl = response.get(Field.get("permanenturl").getResponseName());
        etag = response.get(Field.get("etag").getResponseName());
        displayName = response.get(Field.get("displayname").getResponseName());
        itemName = StringUtil.decodeUrlcompname(response.get(Field.get("urlcompname").getResponseName()));
        // workaround for missing urlcompname in Exchange 2010
        if (itemName == null) {
            itemName = StringUtil.base64ToUrl(itemId.id) + ".EML";
        }
        for (String attributeName : ExchangeSession.CONTACT_ATTRIBUTES) {
            String value = response.get(Field.get(attributeName).getResponseName());
            if (value != null && value.length() > 0) {
                if ("bday".equals(attributeName) || "anniversary".equals(attributeName) || "lastmodified".equals(attributeName) || "datereceived".equals(attributeName)) {
                    value = ewsExchangeSession.convertDateFromExchange(value);
                }
                put(attributeName, value);
            }
        }
    }

    /**
     * @inheritDoc
     */
    EwsContact(EwsExchangeSession ewsExchangeSession, String folderPath, String itemName, Map<String, String> properties, String etag, String noneMatch) {
        super(ewsExchangeSession, folderPath, itemName, properties, etag, noneMatch);
        this.ewsExchangeSession = ewsExchangeSession;
    }

    /**
     * Empty constructor for GalFind
     */
    EwsContact(EwsExchangeSession ewsExchangeSession) {
        super(ewsExchangeSession);
        this.ewsExchangeSession = ewsExchangeSession;
    }

    protected void buildProperties(List<FieldUpdate> updates) {
        for (Map.Entry<String, String> entry : entrySet()) {
            if ("photo".equals(entry.getKey())) {
                updates.add(Field.createFieldUpdate("haspicture", "true"));
            } else if (!entry.getKey().startsWith("email") && !entry.getKey().startsWith("smtpemail")
                    && !entry.getKey().equals("fileas")) {
                updates.add(Field.createFieldUpdate(entry.getKey(), entry.getValue()));
            }
        }
        if (get("fileas") != null) {
            updates.add(Field.createFieldUpdate("fileas", get("fileas")));
        }
        // handle email addresses
        IndexedFieldUpdate emailFieldUpdate = null;
        for (Map.Entry<String, String> entry : entrySet()) {
            if (entry.getKey().startsWith("smtpemail") && entry.getValue() != null) {
                if (emailFieldUpdate == null) {
                    emailFieldUpdate = new IndexedFieldUpdate("EmailAddresses");
                }
                emailFieldUpdate.addFieldValue(Field.createFieldUpdate(entry.getKey(), entry.getValue()));
            }
        }
        if (emailFieldUpdate != null) {
            updates.add(emailFieldUpdate);
        }
    }


    /**
     * Create or update contact
     *
     * @return action result
     * @throws java.io.IOException on error
     */
    public ItemResult createOrUpdate() throws IOException {
        String photo = get("photo");

        ItemResult itemResult = new ItemResult();
        EWSMethod createOrUpdateItemMethod;

        // first try to load existing event
        String currentEtag = null;
        ItemId currentItemId = null;
        FileAttachment currentFileAttachment = null;
        EWSMethod.Item currentItem = ewsExchangeSession.getEwsItem(folderPath, itemName);
        if (currentItem != null) {
            currentItemId = new ItemId(currentItem);
            currentEtag = currentItem.get(Field.get("etag").getResponseName());

            // load current picture
            GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, currentItemId, false);
            getItemMethod.addAdditionalProperty(Field.get("attachments"));
            ewsExchangeSession.executeMethod(getItemMethod);
            EWSMethod.Item item = getItemMethod.getResponseItem();
            if (item != null) {
                currentFileAttachment = item.getAttachmentByName("ContactPicture.jpg");
            }
        }
        if ("*".equals(noneMatch)) {
            // create requested
            if (currentItemId != null) {
                itemResult.status = HttpStatus.SC_PRECONDITION_FAILED;
                return itemResult;
            }
        } else if (etag != null) {
            // update requested
            if (currentItemId == null || !etag.equals(currentEtag)) {
                itemResult.status = HttpStatus.SC_PRECONDITION_FAILED;
                return itemResult;
            }
        }

        List<FieldUpdate> properties = new ArrayList<FieldUpdate>();
        if (currentItemId != null) {
            buildProperties(properties);
            // update
            createOrUpdateItemMethod = new UpdateItemMethod(MessageDisposition.SaveOnly,
                    ConflictResolution.AlwaysOverwrite,
                    SendMeetingInvitationsOrCancellations.SendToNone,
                    currentItemId, properties);
        } else {
            // create
            EWSMethod.Item newItem = new EWSMethod.Item();
            newItem.type = "Contact";
            // force urlcompname on create
            properties.add(Field.createFieldUpdate("urlcompname", exchangeSession.convertItemNameToEML(itemName)));
            buildProperties(properties);
            newItem.setFieldUpdates(properties);
            createOrUpdateItemMethod = new CreateItemMethod(MessageDisposition.SaveOnly, ewsExchangeSession.getFolderId(folderPath), newItem);
        }
        ewsExchangeSession.executeMethod(createOrUpdateItemMethod);

        itemResult.status = createOrUpdateItemMethod.getStatusCode();
        if (itemResult.status == HttpURLConnection.HTTP_OK) {
            //noinspection VariableNotUsedInsideIf
            if (etag == null) {
                itemResult.status = HttpStatus.SC_CREATED;
                LOGGER.debug("Created contact " + getHref());
            } else {
                LOGGER.debug("Updated contact " + getHref());
            }
        } else {
            return itemResult;
        }

        ItemId newItemId = new ItemId(createOrUpdateItemMethod.getResponseItem());

        // disable contact picture handling on Exchange 2007
        if (!exchangeSession.getServerVersion().isExchange2007()) {
            // first delete current picture
            if (currentFileAttachment != null) {
                DeleteAttachmentMethod deleteAttachmentMethod = new DeleteAttachmentMethod(currentFileAttachment.attachmentId);
                ewsExchangeSession.executeMethod(deleteAttachmentMethod);
            }

            if (photo != null) {
                // convert image to jpeg
                byte[] resizedImageBytes = IOUtil.resizeImage(Base64.decodeBase64(photo.getBytes()), 90);

                FileAttachment attachment = new FileAttachment("ContactPicture.jpg", "image/jpeg", new String(Base64.encodeBase64(resizedImageBytes)));
                attachment.setIsContactPhoto(true);

                // update photo attachment
                CreateAttachmentMethod createAttachmentMethod = new CreateAttachmentMethod(newItemId, attachment);
                ewsExchangeSession.executeMethod(createAttachmentMethod);
            }
        }

        GetItemMethod getItemMethod = new GetItemMethod(BaseShape.ID_ONLY, newItemId, false);
        getItemMethod.addAdditionalProperty(Field.get("etag"));
        ewsExchangeSession.executeMethod(getItemMethod);
        itemResult.etag = getItemMethod.getResponseItem().get(Field.get("etag").getResponseName());

        return itemResult;
    }
}
