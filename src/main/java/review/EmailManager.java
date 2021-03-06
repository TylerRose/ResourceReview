package review;

import GUI.MainGUI;

import java.util.ArrayList;
import java.util.Arrays;

/**
 * EmailManager handles the emails and stores them in a list. Ensures that added
 * emails are free of errors marked by the code that generates the email text.
 *
 * @author Tyler Rose
 */
public class EmailManager {

    private static ArrayList<Email> emailList;
    private static EmailManager manager;

    /**
     * Private constructor for Singleton EmailManager. Initializes the list.
     */
    private EmailManager() {
        emailList = new ArrayList<>();
        manager = this;
    }

    /**
     * Get the instance of the EmailManager.
     *
     * @return an EmailManager instance
     */
    public static EmailManager getInstance() {
        if (emailList == null) {
            manager = new EmailManager();
        }
        return manager;
    }

    public static void resetInstance() {
        manager = new EmailManager();
    }

    /**
     * Add an email to the list if there isn't an error
     *
     * @param email the email object to add
     */
    public void addEmail(Email email) {
        if (!email.getTo().toLowerCase().contains("error")
                && !email.getSubject().toLowerCase().contains("error")
                && !email.getBody().toLowerCase().contains("error")) {
            {
                emailList.add(email);
            }
        }
    }

    /**
     * Add an email by To, Subject, and Body fields to the list
     *
     * @param to      the destination email
     * @param subject the email subject
     * @param body    the email body
     */
    public void addEmail(String to, String subject, String body) {
        Email newEmail = new Email(to, subject, body);
        addEmail(newEmail);
    }

    /**
     * Get a list of emails added to the list
     *
     * @return the list of emails
     */
    public ArrayList<Email> getEmails() {
        return emailList;
    }

    public void sendAll() {
        MainGUI.println("Sent " + SendEmail.sentCount + " / " + emailList.size());
        for (Email mail : emailList) {
            if (SendEmail.retryLogin) {
                SendEmail.sendAnEmail(mail);
                MainGUI.replaceLastLog("Sent " + SendEmail.sentCount + " / " + emailList.size());
            }
        }
        if (SendEmail.retryLogin) {
            StringBuilder sheetsRun = new StringBuilder();
            for(Object i : RRMain.sheetList)
            {
                try {
                    sheetsRun.append(i.getClass().getMethod("getSheetName").invoke(i).toString()).append(", ");
                } catch (Exception e){
                    sheetsRun.append("");
                }
            }
            StringBuilder monthsRun = new StringBuilder("");
            for (int i = 0; i < RRMain.sheetList.size(); i++) {
                monthsRun.append(RRMain.sheetList.get(i)+1);
                if(i < RRMain.sheetList.size()-1){
                    monthsRun.append(", ");
                }
            }

            if (SendEmail.tylertest) {
                SendEmail.sendAnEmail(new Email("tylerrose-@outlook.com",
                        (SendEmail.testMode ? "**Test**" : "**Production**") + "Resource Review Emails Finished Sending!",
                        (SendEmail.testMode ? "**Test**" : "**Production**") + "\nYou have finished sending " + SendEmail.sentCount + " emails for " + monthsRun.toString() + " /" + (RRMain.year == -1 ? "TestYear" : RRMain.year) + "."));
            } else {
                SendEmail.sendAnEmail(new Email("resourcereviews@homage.org",
                        (SendEmail.testMode ? "**Test**" : "**Production**") + "Resource Review Emails Finished Sending!",
                        (SendEmail.testMode ? "**Test**" : "**Production**") + "\nYou have finished sending " + SendEmail.sentCount + " emails for " + monthsRun.toString() + " /" + (RRMain.year == -1 ? "TestYear" : RRMain.year) + "."));
            }

        }
    }
}
