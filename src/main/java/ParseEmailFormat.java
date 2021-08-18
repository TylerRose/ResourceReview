
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author trose
 */
public class ParseEmailFormat {

    private final Spreadsheet spdsht;
//    private final Sheet sheet;
//    private final String resourceReviewPath;
    private final String templateLoc;
    private Scanner emailHeadTemplate;
    private Scanner emailBodyTemplate;
    private Scanner emailFooterTemplate;
    private Scanner emailSignatureTemplate = null;
    //private BufferedWriter out;
    private final ErrorTracker errors;

    public ParseEmailFormat(Sheet sheet, String path) throws FileNotFoundException {
        //this.sheet = sheet;
        //resourceReviewPath = path;
        templateLoc = path + "\\EmailFormat\\Email";
        errors = ErrorTracker.getInstance();
        spdsht = Spreadsheet.getInstance();
    }

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

        headerText = ResolveFields(emailHeadTemplate, headerText, currRow);

        //Loop through all their rows to append to the body of the email
        ArrayList<Row> rowLoop = spdsht.getRowsByAdminContactEmail(sheet, email);
        for (Row row : rowLoop) {
            if (spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Complete")).equals("")) {
                emailBodyTemplate = new Scanner(new FileInputStream(new File(templateLoc + "Body.txt")));
                try {
                    //System.out.println(bodyText);
                    //System.out.println("Body x" + rowLoop.size());
                    bodyText = ResolveFields(emailBodyTemplate, bodyText, row);
                    //System.out.println(bodyText);
                    ////test set cell stuffs
                    //System.out.println(spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Data Notes")));
                    //spdsht.appendCellText(spdsht.getCellByRowAndTitle(row, "Dates Contacted"), ", E: " + getDateString());
                    //getDateString();
                    //System.out.println(spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Data Notes")));
                } catch (Exception e) {
                    throw e;
                }
            }
        }

        emailFooterTemplate = new Scanner(new FileInputStream(new File(templateLoc + "Footer.txt")));
        //Write the footer of the email
        footerText = ResolveFields(emailFooterTemplate, footerText, currRow);

        //Write the signature of the email
        emailSignatureTemplate = new Scanner(new FileInputStream(new File(templateLoc + "Signature" + specialistInitials + ".txt")));
        signatureText = ResolveFields(emailSignatureTemplate, signatureText, currRow);

        //Put all the pieces together and write output
        writeFile(headerText.append(bodyText).append(footerText).append(signatureText).toString());
    }

    private StringBuilder ResolveFields(Scanner template, StringBuilder text, Row row) {
        boolean skip = false;
        int start, stop;
        //Loop through every line of the emplate
        while (template.hasNextLine()) {
            text.append(template.nextLine());
            //For each line, loop through every set of brackets
            while (text.toString().contains("[")) {
                start = text.indexOf("[");
                stop = text.indexOf("]");
                String replacementText = "";
                String colTitle = "";
                //Try to insert the ID's information into the email
                try {
                    colTitle = text.substring(start + 1, stop);
                    //System.out.println("Template got column: " + colTitle);
                    replacementText = spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, colTitle));
                    //No error if it was pulling first name, web review.
                    if (replacementText.equals("Administrative Contact First Name")) {
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
                    //String text2 = e.getMessage();
                    //String text3 = Arrays.toString(e.getStackTrace());
                    //System.out.println("\n\n" + text2 + "\n" + text3 + "\n\n");
                    replacementText = "\tERROR - INVALID COLUMN NAME - {" + colTitle + "}\t";
                    String id = spdsht.getCellValue(spdsht.getCellByRowAndTitle(row, "Listing ID"));
                    errors.addError(id.substring(0, id.indexOf(".")), "Invalid Column Name - {" + colTitle + "}");
                }
                //String end = text.substring(stop + 1);
                text.replace(start, stop + 1, replacementText);// = beginning + middle + end;
            }
            //System.out.println("***");
            text.append("\n");
        }
        if (skip) {
            text = null;
        }
        return text;
    }

//    public void parseRow(Spreadsheet spdsht, Row row) throws IOException {
//        String text = "";
//        while (emailTemplate.hasNextLine()) {
//            text += emailTemplate.nextLine();
//            if (text.contains("[")) {
//                int start = text.indexOf("[");
//                int stop = text.indexOf("]");
//                String beginning = text.substring(0, start);
//                String middle;
//                try {
//                    middle = spdsht.getCellByRowAndTitle(row, text.substring(start + 1, stop)).toString();
//                } catch (Exception e) {
//                    middle = "\tERROR - BAD FIELD\t";
//                }
//                String end = text.substring(stop + 1);
//                text = beginning + middle + end;
//            }
//            text += "\n";
//        }
//        writeFile(text);
//    }
    /*
    *Old write file, saves in output as to, subject, body files
     */
//    private void writeFile(String text) throws IOException {
//        String outFilePath = (resourceReviewPath + "\\Outputs\\").concat(text.substring(0, text.indexOf("Subject") - 1)).concat(".txt");
//        File outFile = new File(outFilePath);
//        out = new BufferedWriter(new FileWriter(outFile));
//        out.write(text);
//        out.close();
//    }
    /*
    * new saving emails, saves as email object with the email manager class
     */
    private void writeFile(String text) {
        String[] textArr = text.split("\n");

        String to = textArr[0].substring("to-".length()).trim();
        String subject = textArr[1].substring("subject-".length()).trim();
        StringBuilder body = new StringBuilder();
        for (int i = 2; i < textArr.length; i++) {
            body.append(textArr[i]).append("\n");
        }

        //String to = text.substring(text.toLowerCase().indexOf("to-") + "to-".length(), text.toLowerCase().indexOf("subject")).trim();
        //String subject = text.substring(text.toLowerCase().indexOf("subject-") + "subject-".length(), text.toLowerCase().indexOf("<!doctype html>")).trim();
        // String body = text.substring(text.toLowerCase().indexOf("<!doctype html>"));
        EmailManager.getInstance().addEmail(to, subject, body.toString());
    }
}
