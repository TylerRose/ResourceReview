package review;

import GUI.LoginGUI;
import GUI.MainGUI;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

/**
 * This class handles emails that are being sent. Creates the files needed for
 * Powershell to send the emails and can run the scripts to send them. Errors
 * are moved to a separate folder.
 *
 * @author Tyler Rose
 */
public class SendEmail {

    public static boolean testMode = true;
    private static BufferedWriter out;
    public static int index = 0;
    public static int error = 0;
    private static Process p = null;
    private static String pathCred = null;
    private static String pathEmail = null;
    public static String username = "";
    private static String password = "";
    public static int sentCount = 0;
    public static boolean retryLogin = true;
    public static LoginGUI retryGUI;
    public static Session session = null;
    public static boolean tylertest = false;

    public static void resetSendEmail() {
        testMode = true;
        BufferedWriter out = null;
        index = 0;
        error = 0;
        p = null;
        pathCred = null;
        pathEmail = null;
        username = "";
        password = "";
        sentCount = 0;
        retryLogin = true;
        retryGUI = null;
        session = null;
    }

    /**
     * Add an email to the list of emails to run
     *
     * @param scriptPath the path that the powershell scripts are in
     * @param email      the email object to add
     * @throws IOException Couldn't run the powershell process
     */
    public static void addEmail(File scriptPath, Email email) throws IOException {
        /**
         * Remove powershell functionality
         */
        /*
        //figure out which file is cred/send
        if (pathCred == null || pathEmail == null) {
            if (scriptPath.listFiles()[0].toString().contains("cred")) {
                pathCred = scriptPath.listFiles()[0].toString();
                pathEmail = scriptPath.listFiles()[1].toString();
            } else {
                pathCred = scriptPath.listFiles()[1].toString();
                pathEmail = scriptPath.listFiles()[0].toString();
            }
        }
         */
        //get the email details and write them to file
        String to = email.getTo();
        String subject = email.getSubject();
        String message = email.getBody();
        if (testMode) {
            //test mode only sends to resourcereviews@homage.org
            to = "resourcereviews@homage.org";
            if (tylertest)
                to = "trose@homage.org";
            subject = "Review Test Email - " + subject;
        }
        writeFile(scriptPath + "\\Files\\", "to" + index, to);
        writeFile(scriptPath + "\\Files\\", "subject" + index, subject);
        writeFile(scriptPath + "\\Files\\", "body" + index, message);
        String concat = to.concat(subject).concat(message);
        //Put any errors in a different folder
        if (concat.contains("ERROR") || concat.contains("null")) {
            try {
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y " + scriptPath + "\\Files\\to" + index + ".txt " + scriptPath + "\\..\\Errors\\to" + error + ".txt")).start();
                p.waitFor();
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y " + scriptPath + "\\ps\\Files\\subject" + index + ".txt " + scriptPath + "\\..\\Errors\\subject" + error + ".txt")).start();
                p.waitFor();
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y " + scriptPath + "\\ps\\Files\\body" + index + ".txt " + scriptPath + "\\..\\Errors\\body" + error + ".txt")).start();
                p.waitFor();
                error++;
            } catch (InterruptedException ex) {
                Logger.getLogger(SendEmail.class.getName()).log(Level.SEVERE, null, ex);
            }
        } else {
            index++;
        }
        //Write an error list file to the error folder
        String errorList = "";
        for (String id : ErrorTracker.getInstance().getList()) {
            errorList += id + "\n";
        }
        writeFile(scriptPath + "\\..\\Errors", "errorList.txt", errorList);
    }

    /**
     * Remove powershell functionality
     */
    /**
     * Send emails by calling the powershell script
     *
     * @throws java.io.IOException Couldn't start the powershell process
     * //@deprecated superseded by running powershell after the jar
     */
    /*
    public static void sendAll() throws IOException {
        //Don't send through java, the powershell is run through the batch instead
        p = (new ProcessBuilder("cmd.exe", "/c", "powershell " + pathEmail)).start();
        MainGUI.println("Sending!");
    }
     */

    /**
     * Write text to a new file with a given name and close the file
     *
     * @param file the file name to write to
     * @param text the text to write
     */
    private static void writeFile(String outFilePath, String file, String text) throws IOException {
        outFilePath = outFilePath + file + ".txt";
        File outFile = new File(outFilePath);
        new File(outFilePath).createNewFile();
        out = new BufferedWriter(new FileWriter(outFile));
        out.write(text);
        out.close();
    }

    public static void sendAnEmail(Email email) {
        String host = "smtp.office365.com";
        Properties props = new Properties();

        props.put("mail.transport.protocol", "smtp");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.debug", "false");

        props.put("mail.host", host);
        props.put("mail.smtp.port", "587");

        if (session == null) {
            Authenticate(props);
        }
        try {
            //create a MimeMessage object
            Message message = new MimeMessage(session);

            //set From email field
            //message.setFrom(new InternetAddress("resourcereviews@homage.org"));
            message.setFrom(new InternetAddress(username));

            //set To email field
            if (tylertest) {
                message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("TylerRose-@outlook.com"));
                message.setSubject("Review Test Email - " + email.getSubject());
            } else {
                if (testMode) {
                    message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("resourcereviews@homage.org"));

                    //set email subject field
                    message.setSubject("Review Test Email - " + email.getSubject());
                } else {
                    message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(email.getTo()));
                    message.setRecipients(Message.RecipientType.BCC, InternetAddress.parse("resourcereviews@homage.org"));
                    //set email subject field
                    message.setSubject(email.getSubject());
                }
            }


            //set the content of the email message
            message.setContent(email.getBody(), "text/html");

            //send the email message
            Transport.send(message);
            sentCount++;
        } catch (javax.mail.AuthenticationFailedException e) {
            MainGUI.println("Authentication failed: Please log in again.");
            MainGUI.println("");
            reAuth();
            if (retryLogin) {
                session = null;
                sendAnEmail(email);
            } else {
                throw new RuntimeException("Canceled After Bad Login Credentials");
            }
        } catch (MessagingException e) {
            MainGUI.println("Bad email address or unknown messaging error: Please log in again or cancel and re-launch.");
            MainGUI.println("");
            reAuth();
            if (retryLogin) {
                session = null;
                sendAnEmail(email);
            } else {
                throw new RuntimeException("Canceled After an Unknown Messaging Error");
            }
        }
    }

    private static void reAuth() {
        Thread window = new Thread() {
            @Override
            public void run() {
                SendEmail.retryGUI = new LoginGUI();
                retryGUI.setVisible(true);
                while (retryGUI.isVisible()) ;
            }
        };
        window.start();
        while (true) {
            if (retryGUI != null) {
                if (!retryGUI.isVisible()) {
                    break;
                }
            }
            Source.delay(0.5);
        }
        MainGUI.println("Authenticating...");
        retryGUI.dispose();
        retryGUI = null;
    }

    private static void Authenticate(Properties props) {
        //create the Session object
        session = Session.getInstance(props,
                new javax.mail.Authenticator() {
                    @Override
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });
    }

    public static void setPassword(String pass) {
        password = pass;
    }
}
