/*
 * DavMail POP/IMAP/SMTP/CalDav/LDAP Exchange Gateway
 * Copyright (C) 2009  Mickael Guessant
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
package davmail.smtp;

import davmail.AbstractConnection;
import davmail.BundleMessage;
import davmail.DavGateway;
import davmail.exception.DavMailException;
import davmail.io.DoubleDotInputStream;
import davmail.exchange.ExchangeSessionFactory;
import davmail.ui.tray.DavGatewayTray;

import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.util.SharedByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.Socket;
import java.net.SocketException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.StringTokenizer;

/**
 * Dav Gateway smtp connection implementation
 */
public class SmtpConnection extends AbstractConnection {

    final List<String> recipients = new ArrayList<String>();

    /**
     * Initialize the streams and start the thread.
     *
     * @param clientSocket SMTP client socket
     */
    public SmtpConnection(Socket clientSocket) {
        super(SmtpConnection.class.getSimpleName(), clientSocket, null);
    }

    @Override
    public void run() {

        try {
            sessionFactory.checkConfig();
            sendClient("220 DavMail " + DavGateway.getCurrentVersion() + " SMTP ready at " + new Date());
            for (; ;) {
                String line = readClient();
                // unable to read line, connection closed ?
                if (line == null) {
                    break;
                }

                StringTokenizer tokens = new StringTokenizer(line);
                if (tokens.hasMoreTokens()) {
                    String command = tokens.nextToken();

                    if (state == State.LOGIN) {
                        // AUTH LOGIN, read userName
                        userName = base64Decode(line);
                        sendClient("334 " + base64Encode("Password:"));
                        state = State.PASSWORD;
                    } else if (state == State.PASSWORD) {
                        // AUTH LOGIN, read password
                        password = base64Decode(line);
                        authenticate();
                    } else if ("QUIT".equalsIgnoreCase(command)) {
                        sendClient("221 Closing connection");
                        break;
                    } else if ("NOOP".equalsIgnoreCase(command)) {
                        sendClient("250 OK");
                    } else if ("EHLO".equalsIgnoreCase(command)) {
                        sendClient("250-" + tokens.nextToken());
                        // inform server that AUTH is supported
                        // actually it is mandatory (only way to get credentials)
                        sendClient("250-AUTH LOGIN PLAIN");
                        sendClient("250-8BITMIME");
                        sendClient("250 Hello");
                    } else if ("HELO".equalsIgnoreCase(command)) {
                        sendClient("250 Hello");
                    } else if ("AUTH".equalsIgnoreCase(command)) {
                        handleAuth(tokens);
                    } else if ("MAIL".equalsIgnoreCase(command)) {
                        handleMail();
                    } else if ("RCPT".equalsIgnoreCase(command)) {
                        handleReceipt(line);
                    } else if ("DATA".equalsIgnoreCase(command)) {
                        handleData();
                    } else if ("RSET".equalsIgnoreCase(command)) {
                        handleReset();
                    } else {
                        sendClient("500 Unrecognized command");
                    }
                } else {
                    sendClient("500 Unrecognized command");
                }

                os.flush();
            }

        } catch (SocketException e) {
            DavGatewayTray.debug(new BundleMessage("LOG_CONNECTION_CLOSED"));
        } catch (Exception e) {
            DavGatewayTray.log(e);
            try {
                sendClient("500 " + ((e.getMessage() == null) ? e : e.getMessage()));
            } catch (IOException e2) {
                DavGatewayTray.debug(new BundleMessage("LOG_EXCEPTION_SENDING_ERROR_TO_CLIENT"), e2);
            }
        } finally {
            close();
        }
        DavGatewayTray.resetIcon();
    }

    protected void handleAuth(StringTokenizer tokens) throws IOException {
        if (tokens.hasMoreElements()) {
            String authType = tokens.nextToken();
            if ("PLAIN".equalsIgnoreCase(authType) && tokens.hasMoreElements()) {
                decodeCredentials(tokens.nextToken());
                authenticate();
            } else if ("LOGIN".equalsIgnoreCase(authType)) {
                if (tokens.hasMoreTokens()) {
                    // user name sent on auth line
                    userName = base64Decode(tokens.nextToken());
                    sendClient("334 " + base64Encode("Password:"));
                    state = State.PASSWORD;
                } else {
                    sendClient("334 " + base64Encode("Username:"));
                    state = State.LOGIN;
                }
            } else {
                sendClient("451 Error : unknown authentication type '" + authType + "'");
            }
        } else {
            sendClient("451 Error : authentication type not specified");
        }
    }

    protected void handleMail() throws IOException {
        if (state == State.AUTHENTICATED) {
            state = State.STARTMAIL;
            recipients.clear();
            sendClient("250 Sender OK");
        } else if (state == State.INITIAL) {
            sendClient("503 Authentication required");
        } else {
            state = State.INITIAL;
            sendClient("503 Bad sequence of commands");
        }
    }

    protected void handleReceipt(String line) throws IOException {
        if (state == State.STARTMAIL || state == State.RECIPIENT) {
            if (line.toUpperCase().startsWith("RCPT TO:")) {
                state = State.RECIPIENT;
                try {
                    InternetAddress internetAddress = new InternetAddress(line.substring("RCPT TO:".length()));
                    recipients.add(internetAddress.getAddress());
                } catch (AddressException e) {
                    throw new DavMailException("EXCEPTION_INVALID_RECIPIENT", line);
                }
                sendClient("250 Recipient OK");
            } else {
                sendClient("500 Unrecognized command");
            }

        } else {
            state = State.AUTHENTICATED;
            sendClient("503 Bad sequence of commands");
        }
    }

    protected void handleData() throws IOException {
        if (state == State.RECIPIENT) {
            state = State.MAILDATA;
            sendClient("354 Start mail input; end with <CRLF>.<CRLF>");

            try {
                // read message in buffer
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                DoubleDotInputStream doubleDotInputStream = new DoubleDotInputStream(in);
                int b;
                while ((b = doubleDotInputStream.read()) >= 0) {
                    baos.write(b);
                }
                MimeMessage mimeMessage = new MimeMessage(null, new SharedByteArrayInputStream(baos.toByteArray()));
                session.sendMessage(recipients, mimeMessage);
                state = State.AUTHENTICATED;
                sendClient("250 Queued mail for delivery");
            } catch (Exception e) {
                DavGatewayTray.error(e);
                state = State.AUTHENTICATED;
                sendClient("451 Error : " + e + ' ' + e.getMessage());
            }

        } else {
            state = State.AUTHENTICATED;
            sendClient("503 Bad sequence of commands");
        }
    }

    protected void handleReset() throws IOException {
        recipients.clear();

        if (state == State.STARTMAIL ||
                state == State.RECIPIENT ||
                state == State.MAILDATA ||
                state == State.AUTHENTICATED) {
            state = State.AUTHENTICATED;
        } else {
            state = State.INITIAL;
        }
        sendClient("250 OK Reset");
    }

    /**
     * Create authenticated session with Exchange server
     *
     * @throws IOException on error
     */
    protected void authenticate() throws IOException {
        try {
            session = sessionFactory.getInstance(userName, password);
            sendClient("235 OK Authenticated");
            state = State.AUTHENTICATED;
        } catch (Exception e) {
            DavGatewayTray.error(e);
            String message = e.getMessage();
            if (message == null) {
                message = e.toString();
            }
            message = message.replaceAll("\\n", " ");
            sendClient("554 Authenticated failed " + message);
            state = State.INITIAL;
        }

    }

    /**
     * Decode SMTP credentials
     *
     * @param encodedCredentials smtp encoded credentials
     * @throws IOException if invalid credentials
     */
    protected void decodeCredentials(String encodedCredentials) throws IOException {
        String decodedCredentials = base64Decode(encodedCredentials);
        int startIndex = decodedCredentials.indexOf((char) 0);
        if (startIndex >=0) {
            int endIndex = decodedCredentials.indexOf((char) 0, startIndex+1);
            if (endIndex >=0) {
                userName = decodedCredentials.substring(startIndex+1, endIndex);
                password = decodedCredentials.substring(endIndex + 1);
            } else {
                throw new DavMailException("EXCEPTION_INVALID_CREDENTIALS");
            }
        } else {
            throw new DavMailException("EXCEPTION_INVALID_CREDENTIALS");
        }
    }

}

