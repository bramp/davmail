package davmail.exchange.entity;

import davmail.exchange.ExchangeSession;
import davmail.exchange.VCardWriter;
import org.apache.commons.httpclient.HttpException;
import org.apache.commons.httpclient.URIException;
import org.apache.commons.httpclient.util.URIUtil;
import org.apache.log4j.Logger;

import java.io.IOException;
import java.util.Map;

/**
 * Contact object
 */
public abstract class Contact extends Item {

    private static final Logger LOGGER = Logger.getLogger("davmail.exchange.ExchangeSession");

    protected final ExchangeSession exchangeSession;

    /**
     * @inheritDoc
     */
    public Contact(ExchangeSession exchangeSession, String folderPath, String itemName, Map<String, String> properties, String etag, String noneMatch) {
        super(folderPath, itemName.endsWith(".vcf") ? itemName.substring(0, itemName.length() - 3) + "EML" : itemName, etag, noneMatch);
        this.exchangeSession = exchangeSession;
        this.putAll(properties);
    }

    /**
     * @inheritDoc
     */
    public Contact(ExchangeSession exchangeSession) {
        this.exchangeSession = exchangeSession;
    }

    /**
     * Convert EML extension to vcf.
     *
     * @return item name
     */
    @Override
    public String getName() {
        String name = super.getName();
        if (name.endsWith(".EML")) {
            name = name.substring(0, name.length() - 3) + "vcf";
        }
        return name;
    }

    /**
     * Set contact name
     *
     * @param name contact name
     */
    public void setName(String name) {
        this.itemName = name;
    }

    /**
     * Compute vcard uid from name.
     *
     * @return uid
     * @throws org.apache.commons.httpclient.URIException on error
     */
    protected String getUid() throws URIException {
        String uid = getName();
        int dotIndex = uid.lastIndexOf('.');
        if (dotIndex > 0) {
            uid = uid.substring(0, dotIndex);
        }
        return URIUtil.encodePath(uid);
    }

    @Override
    public String getContentType() {
        return "text/vcard";
    }


    @Override
    public String getBody() throws HttpException {
        // build RFC 2426 VCard from contact information
        VCardWriter writer = new VCardWriter();
        writer.startCard();
        writer.appendProperty("UID", getUid());
        // common name
        writer.appendProperty("FN", get("cn"));
        // RFC 2426: Family Name, Given Name, Additional Names, Honorific Prefixes, and Honorific Suffixes
        writer.appendProperty("N", get("sn"), get("givenName"), get("middlename"), get("personaltitle"), get("namesuffix"));

        writer.appendProperty("TEL;TYPE=cell", get("mobile"));
        writer.appendProperty("TEL;TYPE=work", get("telephoneNumber"));
        writer.appendProperty("TEL;TYPE=home", get("homePhone"));
        writer.appendProperty("TEL;TYPE=fax", get("facsimiletelephonenumber"));
        writer.appendProperty("TEL;TYPE=pager", get("pager"));
        writer.appendProperty("TEL;TYPE=car", get("othermobile"));
        writer.appendProperty("TEL;TYPE=home,fax", get("homefax"));
        writer.appendProperty("TEL;TYPE=isdn", get("internationalisdnnumber"));
        writer.appendProperty("TEL;TYPE=msg", get("otherTelephone"));

        // The structured type value corresponds, in sequence, to the post office box; the extended address;
        // the street address; the locality (e.g., city); the region (e.g., state or province);
        // the postal code; the country name
        writer.appendProperty("ADR;TYPE=home",
                get("homepostofficebox"), null, get("homeStreet"), get("homeCity"), get("homeState"), get("homePostalCode"), get("homeCountry"));
        writer.appendProperty("ADR;TYPE=work",
                get("postofficebox"), get("roomnumber"), get("street"), get("l"), get("st"), get("postalcode"), get("co"));
        writer.appendProperty("ADR;TYPE=other",
                get("otherpostofficebox"), null, get("otherstreet"), get("othercity"), get("otherstate"), get("otherpostalcode"), get("othercountry"));

        writer.appendProperty("EMAIL;TYPE=work", get("smtpemail1"));
        writer.appendProperty("EMAIL;TYPE=home", get("smtpemail2"));
        writer.appendProperty("EMAIL;TYPE=other", get("smtpemail3"));

        writer.appendProperty("ORG", get("o"), get("department"));
        writer.appendProperty("URL;TYPE=work", get("businesshomepage"));
        writer.appendProperty("URL;TYPE=home", get("personalHomePage"));
        writer.appendProperty("TITLE", get("title"));
        writer.appendProperty("NOTE", get("description"));

        writer.appendProperty("CUSTOM1", get("extensionattribute1"));
        writer.appendProperty("CUSTOM2", get("extensionattribute2"));
        writer.appendProperty("CUSTOM3", get("extensionattribute3"));
        writer.appendProperty("CUSTOM4", get("extensionattribute4"));

        writer.appendProperty("ROLE", get("profession"));
        writer.appendProperty("NICKNAME", get("nickname"));
        writer.appendProperty("X-AIM", get("im"));

        writer.appendProperty("BDAY", exchangeSession.convertZuluDateToBday(get("bday")));
        writer.appendProperty("ANNIVERSARY", exchangeSession.convertZuluDateToBday(get("anniversary")));

        String gender = get("gender");
        if ("1".equals(gender)) {
            writer.appendProperty("SEX", "2");
        } else if ("2".equals(gender)) {
            writer.appendProperty("SEX", "1");
        }

        writer.appendProperty("CATEGORIES", get("keywords"));

        writer.appendProperty("FBURL", get("fburl"));

        if ("1".equals(get("private"))) {
            writer.appendProperty("CLASS", "PRIVATE");
        }

        writer.appendProperty("X-ASSISTANT", get("secretarycn"));
        writer.appendProperty("X-MANAGER", get("manager"));
        writer.appendProperty("X-SPOUSE", get("spousecn"));

        writer.appendProperty("REV", get("lastmodified"));

        if ("true".equals(get("haspicture"))) {
            try {
                ContactPhoto contactPhoto = exchangeSession.getContactPhoto(this);
                writer.writeLine("PHOTO;BASE64;TYPE=\"" + contactPhoto.contentType + "\";ENCODING=\"b\":");
                writer.writeLine(contactPhoto.content, true);
            } catch (IOException e) {
                LOGGER.warn("Unable to get photo from contact " + this.get("cn"));
            }
        }

        writer.endCard();
        return writer.toString();
    }

}
