
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Parse through the email format files to create the HTML formatted emails from
 * the columns requested for each admin contact.
 *
 * @author Tyler Rose
 */
public class ParseEmailFormat {

    private final Spreadsheet spdsht;
    private final String templateLoc;
    private Scanner emailHeadTemplate;
    private Scanner emailBodyTemplate;
    private Scanner emailFooterTemplate;
    private Scanner emailSignatureTemplate = null;
    private final ErrorTracker errors;

    /**
     * Constructor initializes the file path and instances for the spreadsheet
     * and error tracker classes
     *
     * @param sheet the sheet with the data
     * @param path the path to the EmailFormat folder
     */
    public ParseEmailFormat(Sheet sheet, String path) {
        templateLoc = path + "\\EmailFormat\\Email";
        errors = ErrorTracker.getInstance();
        spdsht = Spreadsheet.getInstance();
    }

    /**
     *
     * Combine all data for the given email address and put it's Email object
     * together. Searches through the sheet according to the address and uses
     * the email format templates to create the subject and body of the email.
     *
     * @param sheet the sheet to use
     * @param email the email address to send an email to
     * @param specialistInitials the initials of the specialist for the
     * signature file to use
     * @throws IOException One of the email format files couldn't be accessed
     */
    public void parseRowsByEmail(Sheet sheet, String email, String specialistInitials) throws IOException {
        //A row to use for information on the contacts
        Row currRow = spdsht.getRowsByAdminContactEmail(sheet, email).get(0);
        StringBuilder headerText = new StringBuilder();
        StringBuilder bodyText = new StringBuilder();
        StringBuilder footerText = new StringBuilder();
        StringBuilder signatureText = new StringBuilder();
        //Write the header of the email
        int contactNo = spdsht.getCellValue(spdsht.getCellByRowAndTitle(currRow, "Dates Contacted")).split(",").length;
        if (contactNo == 1) {
            emailHeadTemplate = new Scanner(new FileInputStream(new File(templateLoc + "HeadFirst.txt")));
        } else {
            emailHeadTemplate = new Scanner(new FileInputStream(new File(templateLoc + "HeadConsec.txt")));
        }
        //Parse the header of the email
        headerText = ResolveFields(emailHeadTemplate, headerText, currRow, "[", "]");

        //Loop through all their rows to append to the body of the email
        ArrayList<Row> rowLoop = spdsht.getRowsByAdminContactEmail(sheet, email);
        for (Row row : rowLoop) {
            if (spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Complete")).equals("")) {
                emailBodyTemplate = new Scanner(new FileInputStream(new File(templateLoc + "Body.txt")));
                try {
                    bodyText = ResolveFields(emailBodyTemplate, bodyText, row, "[", "]");
                } catch (Exception e) {
                    throw e;
                }
            }
        }
        //Parse the footer of the email
        emailFooterTemplate = new Scanner(new FileInputStream(new File(templateLoc + "Footer.txt")));
        //Write the footer of the email
        footerText = ResolveFields(emailFooterTemplate, footerText, currRow, "[", "]");

        //Parse the signature of the email
        try {
            emailSignatureTemplate = new Scanner(new FileInputStream(new File(templateLoc + "Signature" + specialistInitials + ".txt")));
        } catch (FileNotFoundException ex) {
            System.out.println("The signature file for " + specialistInitials + " couldn't be found!");
            throw ex;
        }
        signatureText = ResolveFields(emailSignatureTemplate, signatureText, currRow, "[", "]");

        //Put all the pieces together and write output
        writeFile(headerText.append(bodyText).append(footerText).append(signatureText).toString());
    }

    /**
     * Replace fields set in the email template with the cell values in the
     * spreadsheet
     *
     * @param template the template scanner to use
     * @param text StringBuilder to add the next line to
     * @param row the row of the contact's data
     * @return the StringBuilder with the output text
     */
    private StringBuilder ResolveFields(Scanner template, StringBuilder text, Row row, String open, String close) {
        boolean skip = false;
        int start, stop;
        //Loop through every line of the emplate
        while (template.hasNextLine()) {
            text.append(template.nextLine());
            //For each line, loop through every set of brackets
            while (text.toString().contains(open)) {
                start = text.indexOf(open);
                stop = text.indexOf(close);
                String replacementText = "";
                String colTitle = "";
                //Try to insert the ID's information into the email
                try {
                    colTitle = text.substring(start + 1, stop);
                    replacementText = spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, colTitle)).toLowerCase();
                    //No error if it was pulling first name, web review.
                    if (replacementText.equals(("Administrative Contact First Name").toLowerCase())) {
                        skip = true;
                    }
                    if (replacementText.equals("")) {
                        replacementText = "\tERROR - EMPTY CELL\t";
                        String id = spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Listing ID"));
                        //Mark as hidden error if the first name was blank
                        if (!skip) {
                            errors.addError(id.substring(0, id.indexOf(".")), "Empty " + colTitle + " Cell");
                        } else {
                            errors.addError(id.substring(0, id.indexOf(".")), "Skipped");
                        }
                    }
                    //Catch null pointer exception from unknown column name
                } catch (NullPointerException e) {
                    replacementText = "\tERROR - INVALID COLUMN NAME - {" + colTitle + "}\t";
                    String id = spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Listing ID"));
                    errors.addError(id.substring(0, id.indexOf(".")), "Invalid Column Name - {" + colTitle + "}");
                }
                text.replace(start, stop + 1, replacementText);// = beginning + middle + end;
            }
            text.append("\n");
        }
        if (skip) {
            text = null;
        }
        return text;
    }

    /**
     * Save emails to an email object
     *
     * @param text The name of the file to write
     */
    private void writeFile(String text) {
        String[] textArr = text.split("\n");

        String to = textArr[0].substring("to-".length()).trim();
        String subject = textArr[1].substring("subject-".length()).trim();
        StringBuilder body = new StringBuilder();
        for (int i = 2; i < textArr.length; i++) {
            body.append(textArr[i]).append("\n");
        }
        EmailManager.getInstance().addEmail(to, subject, body.toString());
    }
}
