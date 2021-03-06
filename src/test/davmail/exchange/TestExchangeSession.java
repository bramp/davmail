package davmail.exchange;

import davmail.Settings;
import davmail.exchange.entity.Folder;
import davmail.http.DavGatewaySSLProtocolSocketFactory;

/**
 *  Test Exchange session
 */
public class TestExchangeSession {

    private TestExchangeSession() {
    }

    /**
     * main method
     * @param argv command line arg
     */
    public static void main(String[] argv) {
        // register custom SSL Socket factory
        int currentArg = 0;
        Settings.setConfigFilePath(argv[currentArg++]);
        Settings.load();

        DavGatewaySSLProtocolSocketFactory.register();

        ExchangeSession session;
        // test auth
        try {
            ExchangeSessionFactory sessionFactory = new ExchangeSessionFactory();
            sessionFactory.checkConfig();
            session = sessionFactory.getInstance(argv[currentArg++], argv[currentArg]);

            Folder folder = session.getFolder("INBOX");
            folder.loadMessages();

            //session.purgeOldestTrashAndSentMessages();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
