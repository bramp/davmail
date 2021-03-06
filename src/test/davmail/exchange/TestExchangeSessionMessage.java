/*
 * DavMail POP/IMAP/SMTP/CalDav/LDAP Exchange Gateway
 * Copyright (C) 2010  Mickael Guessant
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
 */
package davmail.exchange;

import davmail.exchange.entity.Folder;
import davmail.exchange.entity.Message;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import javax.mail.util.SharedByteArrayInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.UUID;

/**
 * Test message handling features.
 */
public class TestExchangeSessionMessage extends AbstractExchangeSessionTestCase {
    static Message message;
    static String messageName;

    public void testCreateMessage() throws IOException, MessagingException {
        session.deleteFolder("testfolder");
        session.createMessageFolder("testfolder");
        MimeMessage mimeMessage = createMimeMessage();
        messageName = UUID.randomUUID().toString()+".EML";
        HashMap<String, String> properties = new HashMap<String, String>();
        properties.put("draft", "0");
        session.createMessage("testfolder", messageName, properties, mimeMessage);
    }

    public void testSearchInbox() throws IOException, MessagingException {
        MessageList messageList = session.searchMessages("INBOX");
        assertNotNull(messageList);
    }

    public void testSearchTrash() throws IOException, MessagingException {
        MessageList messageList = session.searchMessages("Trash");
        assertNotNull(messageList);
    }

    public void testSearchMessage() throws IOException, MessagingException {
        MessageList messageList = session.searchMessages("testfolder");
        assertNotNull(messageList);
        assertEquals(1, messageList.size());
        message = messageList.get(0);
        assertFalse(message.answered);
        assertFalse(message.forwarded);
        assertFalse(message.flagged);
        assertFalse(message.draft);
        assertTrue(message.size > 0);
        assertFalse(message.deleted);
        assertFalse(message.read);
        assertNotNull(message.date);
    }

    public void testFlagMessage() throws IOException, MessagingException {
        Folder testFolder = session.getFolder("testfolder");
        testFolder.loadMessages();
        HashMap<String,String> properties = new HashMap<String,String>();
        properties.put("flagged", "2");
        session.updateMessage(message, properties);

        // refresh folder
        testFolder.loadMessages();
        assertNotNull(testFolder.get(0));
        assertTrue(testFolder.get(0).flagged);
        assertEquals(message.getImapUid(), testFolder.get(0).getImapUid());
    }

    public void testGetContent() throws IOException, MessagingException {
        byte[] content = session.getContent(message);
        assertNotNull(content);
        MimeMessage mimeMessage = new MimeMessage(null, new SharedByteArrayInputStream(content));
        assertTrue(mimeMessage.getHeader("To")[0].indexOf("test@test.local") >= 0);
        assertEquals("Test subject", mimeMessage.getSubject());
        assertEquals("Test message\n", mimeMessage.getContent());
    }

    public void testProcessMessage() throws IOException, MessagingException {
        session.processItem("testfolder", messageName);
    }

    public void testFolderUidNext() throws IOException, MessagingException {
        Folder folder = session.getFolder("testfolder");
        assertTrue(folder.uidNext > 0);
    }

    public void testDeleteMessage() throws IOException {
        session.deleteMessage(message);
        MessageList messageList = session.searchMessages("testfolder");
        assertNotNull(messageList);
        assertEquals(0, messageList.size());
    }

    public void testSpecialMessageCharacter() throws IOException, MessagingException {
        session.deleteFolder("testfolder");
        session.createMessageFolder("testfolder");
        MimeMessage mimeMessage = createMimeMessage();
        messageName = "Special & accenté.EML";
        HashMap<String, String> properties = new HashMap<String, String>();
        properties.put("draft", "0");
        session.createMessage("testfolder", messageName, properties, mimeMessage);
        MessageList messageList = session.searchMessages("testfolder", session.isEqualTo("urlcompname", messageName));
        assertNotNull(messageList);
        assertEquals(1, messageList.size());
    }

    public void testSlashMessageName() throws IOException, MessagingException {
        session.deleteFolder("testfolder");
        session.createMessageFolder("testfolder");
        MimeMessage mimeMessage = createMimeMessage();
        messageName = "test _xF8FF_ slash.EML";
        HashMap<String, String> properties = new HashMap<String, String>();
        properties.put("draft", "0");
        session.createMessage("testfolder", messageName, properties, mimeMessage);
        MessageList messageList = session.searchMessages("testfolder", session.isEqualTo("urlcompname", messageName));
        assertNotNull(messageList);
        assertEquals(1, messageList.size());
    }

    public void testPlusMessageName() throws IOException, MessagingException {
        // fails on Exchange 2003
        session.deleteFolder("testfolder");
        session.createMessageFolder("testfolder");
        MimeMessage mimeMessage = createMimeMessage();
        messageName = "test + plus.EML";
        HashMap<String, String> properties = new HashMap<String, String>();
        properties.put("draft", "0");
        session.createMessage("testfolder", messageName, properties, mimeMessage);
        MessageList messageList = session.searchMessages("testfolder", session.isEqualTo("urlcompname", messageName));
        assertNotNull(messageList);
        assertEquals(1, messageList.size());
    }

    /**
     * Cleanup
     */
    public void testDeleteFolder() throws IOException {
        session.deleteFolder("testfolder");
    }

    public void testSearchAaa() throws IOException, MessagingException {
        MessageList messageList = session.searchMessages("aabb");
        assertNotNull(messageList);
    }

}
