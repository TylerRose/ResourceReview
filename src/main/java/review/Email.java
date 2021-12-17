package review;



/**
 * The Email class holds To, Subject, and Body fields for an email
 *
 * @author Tyler Rose
 */
public class Email {

    private String to;
    private String subject;
    private String body;

    /**
     * Constructor creating a new email object
     *
     * @param to the destination email
     * @param subject the email subject
     * @param body the email body
     */
    public Email(String to, String subject, String body) {
        this.to = to;
        this.subject = subject;
        this.body = body;
    }

    /**
     * Get the To field of the email
     *
     * @return the destination email
     */
    public String getTo() {
        return to;
    }

    /**
     * Set the To field of the email
     *
     * @param to the new destination email
     */
    public void setTo(String to) {
        this.to = to;
    }

    /**
     * Get the Subject field of the email
     *
     * @return the email subject
     */
    public String getSubject() {
        return subject;
    }

    /**
     * Set the Subject field of the email
     *
     * @param subject the new email subject
     */
    public void setSubject(String subject) {
        this.subject = subject;
    }

    /**
     * Get the Body field of the email
     *
     * @return the email body
     */
    public String getBody() {
        return body;
    }

    /**
     * Set the Body field of the email
     *
     * @param body the new email body
     */
    public void setBody(String body) {
        this.body = body;
    }
}
