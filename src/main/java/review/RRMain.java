package review;

import GUI.MainGUI;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author tyler
 */
public class RRMain {

    public static Spreadsheet mySpreadSheet = Spreadsheet.getInstance();
    public static String resourceReviewsPath;
    public static String specialistInitials;
    public static CellWriter writer;
    public static boolean errorsOnly = false;
    public static int year;
    public static Thread workbookSetup;
    public static String excelPath;
    public static String powershellScript;
    public static String fileLocation;
    public static ArrayList<String> specialistList;
    public static ArrayList<String> doneIDs = new ArrayList<>();
    public static int sheetNo;

    public static boolean ReviewSteps(boolean errorsOnly) {
        //Get input for month and year
        //int sheetNo = Integer.parseInt(MainGUI.getInstance().getTxtMonth());//getInput();
        //Give the month to the powershell file to use for confirmation email
        /**
         * Remove Powershell functions
         *
         */
        /*
        String[] monthNames = {"January", "Feburary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
        try {
            new File(powershellScript + "\\lastRun.txt").createNewFile();
            File outFile = new File(powershellScript + "\\lastRun.txt");
            try (BufferedWriter out = new BufferedWriter(new FileWriter(outFile))) {
                SimpleDateFormat formatter = new SimpleDateFormat("MM.dd.yyyy HH:mm:ss");
                Date date = new Date();
                out.write("Script run on " + monthNames[sheetNo] + ", " + (year == -1 ? "TestYear" : year) + " at " + formatter.format(date) + ".");
            }
        } catch (IOException ex) {
            Source.printError("Missing permissions to write to file! (Path: " + powershellScript + "\\lastRun.txt)");
        }
         */
        //Begin processing the spreadsheet
        MainGUI.print("\nProcessing");
        Source.delay(0.5);
        //Define the Spreadsheet and parse location and check it exits
        File docPath = new File(excelPath);
        if (docPath.list().length == 0) {
            Source.printError("Your Excel file couldn't be found at: \n\t" + excelPath + "\nPlease check this file's location and try again.");
            return false;
        }
        //Get the path to the file to parse through it
        fileLocation = excelPath + "\\" + docPath.list()[0];
        try {
            RRMain.mySpreadSheet.setupSpreadsheet(workbookSetup, sheetNo);
        } catch (InterruptedException ex) {
            Source.printError("The worksheet could not be set up.");
        }
        RRMain.writer = CellWriter.getInstance();
        ArrayList<Integer> errorsList = null;
        Sheet sheet = null;
        if (errorsOnly) {
            try {
                //Begin re-running only errored emails by populating the errors list
                errorsList = RunErrorsOnly();
            } catch (IOException ex) {
                Source.printError("The error list wasn't able to be read or couldn't be found! Were there any errors to re-run?");
                return false;
            }
        }
        try {
            //Beginning running the review. If errors is no longer null, it will run the errors only.
            sheet = RRMain.RunResourceReview(sheetNo, errorsList);
        } catch (IllegalArgumentException ex) {
            //Tabs out of order or null cell passed
            Source.printError(ex.getMessage());
        } catch (FileNotFoundException ex) {
            Source.printError("The spreadsheet could not be set up, the file was inaccessable.");
        } catch (RuntimeException ex) {
            Source.printError(ex.getMessage());
        }
        MainGUI.println("Finished processing");
        if (sheet != null) {
            //Generate emails for non-errored entries to be sent by java
            MainGUI.print("\nGenerating emails");
            Source.delay(0.5);
            try {
                sendEmails();
            } catch (IOException ex) {
                Source.printError("Unable to send emails. A file permission error has occured. Please check permissions and try re-running the application.");
            } catch (RuntimeException ex) {
                Source.printError("Unable to send emails: \n" + ex.getMessage() + "\nPlease try re-running the application.");
                return false;
            }
            ErrorTracker errors = reportErrors();
            //write the dates into the contact column
            MainGUI.print("\nUpdating dates for successful entries");
            Source.delay(0.5);
            try {
                RRMain.updateDates(sheet, errors);
                RRMain.updateFormulas(sheet);
            } catch (RuntimeException ex) {
                Source.printError("A runtime exception occured:\n" + ex.getMessage() + "\nPlease contact support with this message and the following information:");
                ex.printStackTrace();
                return false;
            }
            try {
                RRMain.writer.closeWriter();
                MainGUI.println("Finished updating dates.");
            } catch (IOException ex) {
                Source.printError("Unable to save and close the sheet. Dates have not been updated. Make sure the sheet is closed before running.");
                return false;
            }
            //Done :)
            MainGUI.println("\n\n"
                    + "  ____      U  ___ u  _   _   U _____ u \n"
                    + " |  _\"\\      \\/\"_ \\/ | \\ |\"|  \\| ___\"|/ \n"
                    + "/| | | |     | | | |<|  \\| |>  |  _|\"   \n"
                    + "U| |_| |\\.-,_| |_| |U| |\\  |u  | |___   \n"
                    + " |____/ u \\_)-\\___/  |_| \\_|   |_____|  \n"
                    + "  |||_         \\\\    ||   \\\\,-.<<   >>  \n"
                    + " (__)_)       (__)   (_\")  (_/(__) (__) \n"
                    + "                                        \n"
                    + "                                        \n");
        } else {
            MainGUI.println("The sheet to run wasn't identified or another error has occured. Please read the information above.");
            MainGUI.println("Nothing has been run.");
            return false;
        }
        return true;
    }

    /**
     * Run the review on the items with errors only by populating the errored ID
     * list
     */
    private static ArrayList<Integer> RunErrorsOnly() throws IOException {
        ArrayList<Integer> out = new ArrayList<>();
        FileInputStream list = new FileInputStream(new File(RRMain.resourceReviewsPath + "\\Errors\\errorList.txt"));
        Scanner in = new Scanner(list);
        while (in.hasNextLine()) {
            out.add(Integer.parseInt(in.nextLine()));
        }
        return out;
    }

    /**
     * This function checks every line in the sheet and for each unique email
     * address it finds it calls ParseEmailFormat's function that will combine
     * that contact's information into one singular email.
     *
     * @param sheetNo the sheet number to run
     * @return the sheet that was run
     * @throws IllegalArgumentException Invalid tab position of the given tab to
     * run
     * @throws FileNotFoundException ParseEmailFormat couldn't open the file at
     * the given path
     */
    public static Sheet RunResourceReview(int sheetNo, ArrayList<Integer> errors) throws FileNotFoundException, RuntimeException {
        ParseEmailFormat parse = new ParseEmailFormat(mySpreadSheet.getSheet(sheetNo), resourceReviewsPath);
        //Parse through the email addresses, combining all agencies per address before moving to the next
        Sheet sheet = mySpreadSheet.getSheet(sheetNo);
        //Ensure tabs are in the correct order before running the sheet
        CheckTabOrder(sheet, sheetNo);
        //Ensure the formulas for each column are accurate
        //CheckFormulaIntegrity(sheet);
        //Ensure the listing name matches the listing text for the URL hyperlink
        CheckListingNameIntegrity(sheet);

        String prevEmail = "----";
        String currEmail = "";
        ArrayList<String> done = new ArrayList<>();
        //set an absolute maximum of 10k lines that will be processed
        int maxRow = 10000;
        int curRow = 0;
        //For each row in the sheet, get the unprocessed, completed email addresses and combine their information
        for (Row row : sheet) {
            try {
                //Check that the current row has an Agency ID and isn't complete
                Cell cell = mySpreadSheet.getCellByRowAndTitle(row, "Agency ID");
                Cell completed = mySpreadSheet.getCellByRowAndTitle(row, "Complete");
                //If the cell isn't null, completed is empty, and no errors to use
                if (cell != null && mySpreadSheet.getCellValue(completed).equals("") && errors == null) {
                    currEmail = (mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Administrative Contact Email"))).toLowerCase();
                    //If it is a new unique email, compile all this emails data
                    if (currEmail.length() > 3 && !currEmail.equals(prevEmail) && !done.contains(currEmail)) {
                        done.add(currEmail);
                        //Get all other lines with the current email address
                        parse.parseRowsByEmail(mySpreadSheet.getSheet(sheetNo), currEmail, specialistInitials);
                        prevEmail = currEmail;
                    }
                    //else if the cell isn't null, completed is empty, and there is an error list to use
                } else if (cell != null && mySpreadSheet.getCellValue(completed).equals("") && errors != null) {
                    //Only if the error list contains the row's listing ID is the listing pushed to the parser.
                    if (errors.contains(Integer.parseInt(mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing ID"))))) {
                        currEmail = (mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Administrative Contact Email"))).toLowerCase();
                        //If it is a new unique email, compile all this emails data
                        if (currEmail.length() > 3 && !currEmail.equals(prevEmail) && !done.contains(currEmail)) {
                            done.add(currEmail);
                            //Get all other lines with the current email address
                            parse.parseRowsByEmail(mySpreadSheet.getSheet(sheetNo), currEmail, specialistInitials);
                            prevEmail = currEmail;
                        }
                    }
                }
            } catch (IOException | RuntimeException e) {
            }
            curRow++;
            //Catch a runaway loop that looks through too many lines
            if (curRow >= maxRow) {
                return sheet;
            }
        }
        return sheet;
    }

    /**
     * Update formulas to show the evaluated value
     *
     * @param sheet the sheet number being run
     * @throws NullPointerException A cell didn't exist
     */
    public static void updateFormulas(Sheet sheet) throws NullPointerException {
        for (Row row : sheet) {
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Contact No"));
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Next Email Ordinal"));
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Latest Contact"));
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Next Contact No"));
        }
    }

    /**
     * Set the formulas in the formula cells to ensure they are correct
     *
     * @param sheet the sheet that will be run
     */
    public static void CheckFormulaIntegrity(Sheet sheet) {
        //Define what the formulas should be and set them in the sheet
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            if (mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing ID")).equals("")) {
                break;
            }
            String rowNum = row.getRowNum() + 1 + "";
            String contactNoFormula = "=IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\":\",\"\"))),0)";
            String nextEmailOrdinalFormula = "=IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\"E\",\"\"))),0)+1 &IF(MOD(ABS(IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\"E\",\"\"))),0)),100)+1>=4,\"th\",CHOOSE(MOD(ABS(IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\"E\",\"\"))),0)+1),10)+1,\"th\",\"st\",\"nd\",\"rd\"))";
            String lastContactFormula = "=IFERROR(RIGHT(W" + rowNum + ",LEN(W" + rowNum + ")-1-FIND(\"@\",SUBSTITUTE(W" + rowNum + ",\",\",\"@\",LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\",\",\"\"))),1)),W" + rowNum + ")";
            writer.setCellText(mySpreadSheet.getCellByRowAndTitle(row, "Contact No"), contactNoFormula);
            //MainGUI.println("Cell before:" + mySpreadSheet.getCellByRowAndTitle(row, "Next Email Ordinal"));
            writer.setCellText(mySpreadSheet.getCellByRowAndTitle(row, "Next Email Ordinal"), nextEmailOrdinalFormula);
            //MainGUI.println("Cell after:" + mySpreadSheet.getCellByRowAndTitle(row, "Next Email Ordinal"));
            writer.setCellText(mySpreadSheet.getCellByRowAndTitle(row, "Latest Contact"), lastContactFormula);
            //MainGUI.println("Formula updated for row #"+rowNum);
        }
        updateFormulas(sheet);
    }

    public static void CheckListingNameIntegrity(Sheet sheet) throws RuntimeException {
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            if (mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing ID")).equals("")) {
                break;
            }
            if (!mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing Name")).toLowerCase().equals(mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Consumer URL Text")).toLowerCase())) {
                String ID = mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing ID"));
                ID = ID.substring(0, ID.indexOf("."));
                throw new RuntimeException("There was a listing where the Listing Name doesn't match the URL Text! - Listing " + ID + ": " + mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing Name")));
            }

        }
    }

    /**
     * Check that the number in the tab name matches the actual tab number of
     * the worksheet. This protects against the requested month not matching the
     * tab number provided by the worksheet and prevents the wrong month from
     * being run accidentally. Throws an exception if the tabs are out of order.
     *
     * @param sheet the sheet that will be run
     * @param sheetNo the month number requested
     * @throws IllegalArgumentException the provided requested sheet number does
     * not match the name of the tab on the sheet.
     */
    public static void CheckTabOrder(Sheet sheet, int sheetNo) throws IllegalArgumentException {
        //clean up the tab name and set the expected name string
        String fullTabName = sheet.getSheetName().toLowerCase().replace(" ", "").replace("_", "").replace("-", "");
        String lookingFor = "tab" + (sheetNo + 1);
        char nextChar;
        //Check for cases like looking for tab1 but the tab is tab10, should not run
        if (fullTabName.length() > fullTabName.indexOf(lookingFor) + lookingFor.length() + 1) {
            nextChar = fullTabName.charAt(fullTabName.indexOf(lookingFor) + lookingFor.length() + 1);
        } else {
            nextChar = 'a';
        }
        //throw exception if the tab name matches the expected name
        if (!(fullTabName).contains(lookingFor) || Character.isDigit(nextChar)) {
            throw new IllegalArgumentException("The tabs were out of order! Please re-order the tabs and run again.");
        }
    }

    /**
     * Update the Dates Contacted field of any lines that ran through the script
     *
     * @param sheet the sheet that was run
     * @param errors the error tracker instance holding which entries had errors
     * @throws RuntimeException
     */
    public static void updateDates(Sheet sheet, ErrorTracker errors) throws RuntimeException {
        Cell currCell;
        String currCellText;
        int columnNum = mySpreadSheet.getCellByRowAndTitle(sheet.getRow(0), "Dates Contacted").getColumnIndex();
        int maxRow = 10000;
        int curRow = 0;
        //go through all rows with Listing IDs
        for (Row r : mySpreadSheet.getRowsByColumnName(sheet, "Listing ID")) {
            try {
                currCell = mySpreadSheet.getCellByRowAndTitle(r, "Dates Contacted");
                currCellText = mySpreadSheet.getCellValue(currCell);
                if (!currCellText.equals("Dates Contacted")) {
                    //if not marked as an error and not completed
                    if (!errors.checkForError(mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(r, "Listing ID")))
                            && mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(r, "Complete")).equals("")) {
                        try {
                            if (mySpreadSheet.getCellValue(currCell).length() > 3) {
                                writer.appendCellText(currCell, ", E: " + Spreadsheet.getDateString());
                            } else {
                                writer.appendCellText(currCell, "E: " + Spreadsheet.getDateString());
                            }
                        } catch (RuntimeException ex) {
                            //Null cell - create new cell
                            if (mySpreadSheet.getCellValue(currCell).length() > 3) {
                                writer.appendCellText(r.createCell(mySpreadSheet.getCellByRowAndTitle(r, "Dates Contacted").getColumnIndex()), ", E: " + Spreadsheet.getDateString());
                            } else {
                                writer.appendCellText(r.createCell(mySpreadSheet.getCellByRowAndTitle(r, "Dates Contacted").getColumnIndex()), "E: " + Spreadsheet.getDateString());
                            }
                        }
                    }
                }
            } catch (NullPointerException e) {
                Cell newCell = r.createCell(columnNum);
                newCell.setCellValue("");
                currCell = newCell;
                writer.appendCellText(currCell, "E: " + Spreadsheet.getDateString());
            }
            curRow++;
            if (curRow >= maxRow) {
                MainGUI.println("Hit 10k rows...");
                return;
            }
        }
    }

    /**
     * Call the function to add run entries to the email list and run the emails
     *
     * @throws IOException files couldn't be created for this email
     */
    private static void sendEmails() throws IOException, RuntimeException {
        File scriptLoc = new File(RRMain.resourceReviewsPath + "ps\\");
        EmailManager eManager = EmailManager.getInstance();
        ArrayList<Email> emailList = eManager.getEmails();
        for (Email e : emailList) {
            SendEmail.addEmail(scriptLoc, e);
        }
        eManager.sendAll();
    }

    /**
     * Send any errored items to the error tracker to record and display which
     * entries had errors.
     *
     * @return the error tracker instance
     */
    private static ErrorTracker reportErrors() {
        ErrorTracker errors = ErrorTracker.getInstance();
        MainGUI.println("Finished sending with " + errors.getNumOfErrors() + " errors.");
        if (errors.getNumOfErrors() > 0) {
            MainGUI.println("The following Listing ID's were not notified:");
            errors.printErrorList();
            MainGUI.println("Please check the error folder.");
        }
        return errors;
    }

}
