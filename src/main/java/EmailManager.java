
import java.util.ArrayList;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author trose
 */
public class EmailManager {

    private static ArrayList<Email> emailList;
    private static EmailManager manager;

    private EmailManager() {
        emailList = new ArrayList<>();
        manager = this;
    }

    public static EmailManager getInstance() {
        if (emailList == null) {
            return new EmailManager();
        } else {
            return manager;
        }
    }

    public void addEmail(Email email) {
        if (!email.getTo().toLowerCase().contains("error")
                && !email.getSubject().toLowerCase().contains("error")
                && !email.getBody().toLowerCase().contains("error")) {
            {
                emailList.add(email);
            }
        }
    }

    public void addEmail(String to, String subject, String body) {
        Email newEmail = new Email(to, subject, body);
        addEmail(newEmail);
    }
    
    public ArrayList<Email> getEmails(){
        return emailList;
    }

}
