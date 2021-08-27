
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

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
    private static int index = 0;
    private static int error = 0;
    private static Process p = null;
    private static String pathCred = null;
    private static String pathEmail = null;

    /**
     * Add an email to the list of emails to run
     *
     * @param scriptPath the path that the powershell scripts are in
     * @param email the email object to add
     * @throws IOException Couldn't run the powershell process
     */
    public static void addEmail(File scriptPath, Email email) throws IOException {
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
        //get the email details and write them to file
        String to = email.getTo();
        String subject = email.getSubject();
        String message = email.getBody();
        if (testMode) {
            //test mode only sends to resourcereviews@homage.org
            to = "resourcereviews@homage.org";
            //to = "trose@homage.org";
            subject = "Review Test Email - " + subject;
        }
        writeFile("to" + index, to);
        writeFile("subject" + index, subject);
        writeFile("body" + index, message);
        String concat = to.concat(subject).concat(message);
        //Put any errors in a different folder
        if (concat.contains("ERROR") || concat.contains("null")) {
            try {
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y C:\\ResourceReviews\\ps\\Files\\to" + index + ".txt " + "C:\\ResourceReviews\\Errors\\to" + error + ".txt")).start();
                p.waitFor();
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y C:\\ResourceReviews\\ps\\Files\\subject" + index + ".txt " + "C:\\ResourceReviews\\Errors\\subject" + error + ".txt")).start();
                p.waitFor();
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y C:\\ResourceReviews\\ps\\Files\\body" + index + ".txt " + "C:\\ResourceReviews\\Errors\\body" + error + ".txt")).start();
                p.waitFor();
                error++;
            } catch (InterruptedException ex) {
                Logger.getLogger(SendEmail.class.getName()).log(Level.SEVERE, null, ex);
            }
        } else {
            index++;
        }
    }

    /**
     * Send emails by calling the powershell script
     *
     * @throws java.io.IOException Couldn't start the powershell process
     * @deprecated superseded by running powershell after the jar
     */
    public static void sendAll() throws IOException {
        //Don't send through java, the powershell is run through the batch instead
        p = (new ProcessBuilder("cmd.exe", "/c", "powershell " + pathEmail)).start();
        System.out.println("Sending!");
    }

    /**
     * Write text to a new file with a given name and close the file
     *
     * @param file the file name to write to
     * @param text the text to write
     */
    private static void writeFile(String file, String text) {
        String outFilePath = "C:\\ResourceReviewsAutomation\\ps\\Files\\" + file + ".txt";
        File outFile = new File(outFilePath);
        try {
            new File(outFilePath).createNewFile();
            out = new BufferedWriter(new FileWriter(outFile));
            out.write(text);
            out.close();
        } catch (IOException ex) {
            Logger.getLogger(SendEmail.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
