
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class SendEmail {

    public static boolean testMode = true;
    private static BufferedWriter out;
    private static int index = 0;
    private static int error = 0;
    private static Process p = null;
    private static String pathCred = null;
    private static String pathEmail = null;

    public static void addEmail(File scriptPath, String from, Email email) throws IOException {
        if (pathCred == null || pathEmail == null) {
            if (scriptPath.listFiles()[0].toString().contains("cred")) {
                pathCred = scriptPath.listFiles()[0].toString();
                pathEmail = scriptPath.listFiles()[1].toString();
            } else {
                pathCred = scriptPath.listFiles()[1].toString();
                pathEmail = scriptPath.listFiles()[0].toString();
            }
        }

        String to = email.getTo();
        String subject = email.getSubject();
        String message = email.getBody();
        if (testMode) {
            to = "resourcereviews@homage.org";
            //to = "trose@homage.org";
            subject = "Review Test Email - " + subject;
        }
        writeFile("to" + index, to);
        writeFile("subject" + index, subject);
        writeFile("body" + index, message);
        String concat = to.concat(subject).concat(message);

        if (concat.contains("ERROR") || concat.contains("null")) {

            try {
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y C:\\ResourceReviews\\ps\\Files\\to" + index + ".txt " + "C:\\ResourceReviews\\Errors\\to" + error + ".txt")).start();
                p.waitFor();
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y C:\\ResourceReviews\\ps\\Files\\subject" + index + ".txt " + "C:\\ResourceReviews\\Errors\\subject" + error + ".txt")).start();
                p.waitFor();
                p = (new ProcessBuilder("cmd.exe", "/c", "move /Y C:\\ResourceReviews\\ps\\Files\\body" + index + ".txt " + "C:\\ResourceReviews\\Errors\\body" + error + ".txt")).start();
                p.waitFor();
                error++;
                //System.out.println(concat);
            } catch (InterruptedException ex) {
                Logger.getLogger(SendEmail.class.getName()).log(Level.SEVERE, null, ex);
            }
        } else {
            index++;
            //System.out.println(concat);
        }
    }

    public static void sendAll() {
        //Don't send through java, the powershell is run through the batch instead
        /*
        try {
            p = (new ProcessBuilder("cmd.exe", "/c", "powershell " + pathEmail)).start();
            System.out.println("Sending!");
        } catch (IOException ex) {
            Logger.getLogger(SendEmail.class.getName()).log(Level.SEVERE, null, ex);
        }
         */
    }

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
