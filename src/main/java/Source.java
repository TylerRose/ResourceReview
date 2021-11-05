
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.*;

/**
 * Main file of RR that handles the program flow, major steps, and error
 * handling
 *
 * @author Tyler Rose
 */
public class Source {

    private static Spreadsheet mySpreadSheet = Spreadsheet.getInstance();
    private static String resourceReviewsPath;
    private static String excelPath;
    private static String powershellScript;
    private static String specialistInitials;
    private static ArrayList<String> specialistList;
    private static CellWriter writer;
    private static int year;
    private static Thread workbookSetup;

    ArrayList<String> doneIDs = new ArrayList<>();

    private static String fileLocation;// = "C:\\Excel\\Test Excel Sheet-2021.xlsx";
    //private static FileInputStream file;
    //private static Workbook workbook;
    private static final Scanner in = new Scanner(System.in);

    /**
     * Main functions, handles the execution order of the RR automation. No
     * arguments runs in test mode. Set test mode with test=true or test=false.
     * Add "errors" after the test mode to run only errors.
     *
     * @param args command line arguments to enable test mode or run only errors
     */
    public static void main(String[] args) {
        //handle the arguments and enable test/error modes
        boolean errorsOnly = false;
        if (args.length == 0) {
            //set default argument values
            SendEmail.testMode = true;
        } else {
            //Set test mode (send all internal) if first argument is test= true/false
            switch (args[0].toLowerCase()) {
                case "test=true":
                    SendEmail.testMode = true;
                    break;
                case "test=false":
                    //if test mode is off, run a confirmation input step
                    SendEmail.testMode = false;
                    do {
                        System.out.print("Running and sending to RR contacts. **If re-sending errors** the listing IDs will be emailed using the data of the month/year you are about to select.\nType 'yes' to confirm production run: ");
                    } while (!in.nextLine().toLowerCase().contains("yes"));
                    System.out.println("");
                    break;
                default:
                    SendEmail.testMode = true;
                    break;
            }
            if (args.length > 1) {
                switch (args[1].toLowerCase()) {
                    case "errors":
                        errorsOnly = true;
                        break;
                }
            }
        }

        //Initialize vars and files
        try {
            initialization();
        } catch (RuntimeException e) {
            System.out.println("ERROR: " + e.getMessage());
        }

        //Get input for month and year
        int sheetNo = getInput();
        //Give the month to the powershell file to use for confirmation email
        String[] monthNames = {"January", "Feburary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
        try {
            new File(powershellScript + "\\lastRun.txt").createNewFile();
            File outFile = new File(powershellScript + "\\lastRun.txt");
            try (BufferedWriter out = new BufferedWriter(new FileWriter(outFile))) {
                SimpleDateFormat formatter = new SimpleDateFormat("MM.dd.yyyy HH:mm:ss");
                Date date = new Date();
                out.write("Script run on " + monthNames[sheetNo] + ", " + year + " at " + formatter.format(date) + ".");
            }
        } catch (IOException ex) {
            System.out.println("ERROR: Missing permissions to write to file! (Path: " + powershellScript + "\\lastRun.txt)");
        }

        //Begin processing the spreadsheet
        System.out.print("\nProcessing");
        delay(3);
        //Define the Spreadsheet and parse location and check it exits
        File docPath = new File(excelPath);
        if (docPath.list().length == 0) {
            System.out.println("Your Excel file couldn't be found at: ");
            System.out.println("\t" + excelPath);
            System.out.println("Please check this file's location and try again.");
            return;
        }
        //Get the path to the file to parse through it
        fileLocation = excelPath + "\\" + docPath.list()[0];
        try {
            mySpreadSheet.setupSpreadsheet(workbookSetup, sheetNo);
        } catch (InterruptedException ex) {
            System.out.println("ERROR: The worksheet could not be set up.");
        }
        writer = CellWriter.getInstance();
        ArrayList<Integer> errorsList = null;
        Sheet sheet = null;
        if (errorsOnly) {
            try {
                //Begin re-running only errored emails by populating the errors list
                errorsList = RunErrorsOnly();
            } catch (IOException ex) {
                System.out.println("ERROR: The error list wasn't able to be read or couldn't be found! Were there any errors to re-run?");
                return;
            }
        }
        try {
            //Beginning running the review. If errors is no longer null, it will run the errors only.
            sheet = RunResourceReview(sheetNo, errorsList);
        } catch (IllegalArgumentException ex) {
            System.out.println("ERROR: The tabs were out of order! Please re-order the tabs and run again.");
        } catch (FileNotFoundException ex) {
            System.out.println("ERROR: The spreadsheet could not be set up, the file was inaccessable.");
        }
        
        System.out.println("Finished processing");

        if (sheet != null) {
            //Generate emails for non-errored entries to be sent by powershell
            System.out.print("\nGenerating emails");
            delay(3);
            try {
                sendEmails();
            } catch (IOException ex) {
                System.out.println("ERROR: Unable to start powershell processes. Please run manually in the PS folder.");
            }

            ErrorTracker errors = reportErrors();

            //write the dates into the contact column
            System.out.print("\nUpdating dates for successfull entries");
            delay(3);
            try {
                updateDates(sheet, errors);
                updateFormulas(sheet);
            } catch (RuntimeException ex) {
                System.out.println("A runtime exception occured:\n" + ex.getMessage());
                System.out.println("Please contact support with this message and the following information:");
                ex.printStackTrace();
            }
            try {
                writer.closeWriter();
                System.out.println("Finished updating dates.");
            } catch (IOException ex) {
                System.out.println("ERROR: Unable to save and close the sheet. Dates have not been updated. Make sure the sheet is closed before running.");
            }

            //Done :)
            System.out.println("\n\n"
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
            System.out.println("The sheet to run wasn't identified or another error has occured. Please read the information above.");
            System.out.println("Nothing has been run.");
        }
    }

    /**
     * Update formulas to show the evaluated value
     *
     * @param sheet the sheet number being run
     * @throws NullPointerException A cell didn't exist
     */
    private static void updateFormulas(Sheet sheet) throws NullPointerException {
        for (Row row : sheet) {
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Contact No"));
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Next Email Ordinal"));
            writer.refreshCell(mySpreadSheet.getCellByRowAndTitle(row, "Latest Contact"));
        }
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
    private static Sheet RunResourceReview(int sheetNo, ArrayList<Integer> errors) throws IllegalArgumentException, FileNotFoundException, RuntimeException {
        ParseEmailFormat parse = new ParseEmailFormat(mySpreadSheet.getSheet(sheetNo), resourceReviewsPath);
        //Parse through the email addresses, combining all agencies per address before moving to the next
        Sheet sheet = mySpreadSheet.getSheet(sheetNo);
        //Ensure tabs are in the correct order before running the sheet
        CheckTabOrder(sheet, sheetNo);
        //Ensure the formulas for each column are accurate
        CheckFormulaIntegrity(sheet);
        //Ensure the listing name matches the listing text for the URL hyperlink
        //CheckListingNameIntegrity(sheet);

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
    private static void CheckTabOrder(Sheet sheet, int sheetNo) throws IllegalArgumentException {
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
            System.out.println("Tabs our of order!");
            throw new IllegalArgumentException("Tabs out of order!");
        }
    }

    /**
     * Set the formulas in the formula cells to ensure they are correct
     *
     * @param sheet the sheet that will be run
     */
    private static void CheckFormulaIntegrity(Sheet sheet) {
        //Define what the formulas should be and set them in the sheet
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            String rowNum = row.getRowNum() + "";
            String contactNoFormula = "=IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\":\",\"\"))),0)";
            String nextEmailOrdinalFormula = "=IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\"E\",\"\"))),0)+1 &IF(MOD(ABS(IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\"E\",\"\"))),0)),100)+1>=4,\"th\",CHOOSE(MOD(ABS(IF(LEN(W" + rowNum + ")>1,SUM(LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\"E\",\"\"))),0)+1),10)+1,\"th\",\"st\",\"nd\",\"rd\"))";
            String lastContactFormula = "=IFERROR(RIGHT(W" + rowNum + ",LEN(W" + rowNum + ")-1-FIND(\"@\",SUBSTITUTE(W" + rowNum + ",\",\",\"@\",LEN(W" + rowNum + ")-LEN(SUBSTITUTE(W" + rowNum + ",\",\",\"\"))),1)),W" + rowNum + ")";
            writer.setCellText(mySpreadSheet.getCellByRowAndTitle(row, "Contact No"), contactNoFormula);
            writer.setCellText(mySpreadSheet.getCellByRowAndTitle(row, "Next Email Ordinal"), nextEmailOrdinalFormula);
            writer.setCellText(mySpreadSheet.getCellByRowAndTitle(row, "Latest Contact"), lastContactFormula);
        }
        updateFormulas(sheet);
    }

    private static void CheckListingNameIntegrity(Sheet sheet) throws RuntimeException {
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            if (!mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing Name")).equals(mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Consumer URL Text")))) {
                throw new RuntimeException("There was a listing where the Listing Name doesn't match the URL Text! - Listing: " + mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing Name")));
            }

        }
    }

    /**
     * Get year, month, and specialist information to know what is being run and
     * where it is located
     *
     * @return the sheet number that was provided
     */
    private static int getInput() {
        int sheetNo;
        System.out.println("Please enter the year: ");
        year = in.nextInt();
        excelPath = resourceReviewsPath + "Excel\\" + year + "\\";
        //set up the sheet in a separate thread to load data while getting input
        workbookSetup = new Thread() {
            @Override
            public void run() {
                try {
                    mySpreadSheet.setupWorkbook(excelPath + "\\" + new File(excelPath).list()[0]);
                } catch (IOException ex) {
                    System.out.println("ERROR: The workbook could not be set up, the file was inaccessable.");
                }
            }
        };
        workbookSetup.start();

        //Create the path if it doesn't exist
        new File(excelPath).mkdirs();
        //Get month and specialist details
        System.out.println("Please enter the Month number: ");
        sheetNo = in.nextInt() - 1;
        System.out.print("Please select the specialist:");
        int i = 1;
        for (String s : specialistList) {
            System.out.print(" (" + i++ + ") " + s);
        }
        System.out.println("");
        int selected = in.nextInt();
        String name = specialistList.get(selected - 1);
        specialistInitials = name.substring(0, 1).concat(name.substring(name.indexOf(" ") + 1, name.indexOf(" ") + 2));
        return sheetNo;
    }

    /**
     * Send any errored items to the error tracker to record and display which
     * entries had errors.
     *
     * @return the error tracker instance
     */
    private static ErrorTracker reportErrors() {
        ErrorTracker errors = ErrorTracker.getInstance();
        System.out.println("Finished sending with " + errors.getNumOfErrors() + " errors.");
        if (errors.getNumOfErrors() > 0) {
            System.out.println("The following Listing ID's were not notified:");
            errors.printErrorList();
            System.out.println("Please check the error folder.");
        }
        return errors;
    }

    /**
     * Update the Dates Contacted field of any lines that ran through the script
     *
     * @param sheet the sheet that was run
     * @param errors the error tracker instance holding which entries had errors
     * @throws RuntimeException
     */
    private static void updateDates(Sheet sheet, ErrorTracker errors) throws RuntimeException {
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
                System.out.println("Hit 10k rows...");
                return;
            }
        }
    }

    /**
     * Call the function to add run entries to the email list and run the emails
     *
     * @throws IOException files couldn't be created for this email
     */
    private static void sendEmails() throws IOException {
        File scriptLoc = new File(resourceReviewsPath + "ps\\");
        EmailManager eManager = EmailManager.getInstance();
        ArrayList<Email> emailList = eManager.getEmails();
        for (Email e : emailList) {
            SendEmail.addEmail(scriptLoc, e);
        }
        //Not sending via Java
        //SendEmail.sendAll();
    }

    /**
     * Initialize variables and files needed by the program
     */
    private static void initialization() {
        FileInputStream specialists = null;
        specialistList = new ArrayList<>();
        resourceReviewsPath = "C:\\ResourceReviewsAutomation\\";
        powershellScript = "C:\\ResourceReviewsAutomation\\ps\\";
        try {
            new File(resourceReviewsPath + "Specialists.txt").createNewFile();
            specialists = new FileInputStream(new File(resourceReviewsPath + "Specialists.txt"));
        } catch (IOException e) {
            System.out.println("Could not access specialists file.");
        }
        if (specialists != null) {
            Scanner specialistIn = new Scanner(specialists);
            while (specialistIn.hasNextLine()) {
                specialistList.add(specialistIn.nextLine());
            }
        }
        if (specialistList.isEmpty()) {
            throw new RuntimeException("Specialist list must not be empty");
        }
        //Create any paths that didn't exist
        new File(resourceReviewsPath).mkdirs();
        new File(powershellScript).mkdirs();
    }

    /**
     * Run the review on the items with errors only by populating the errored ID
     * list
     */
    private static ArrayList<Integer> RunErrorsOnly() throws IOException {
        ArrayList<Integer> out = new ArrayList<>();
        FileInputStream list = new FileInputStream(new File(resourceReviewsPath + "\\Errors\\errorList.txt"));
        Scanner in = new Scanner(list);
        while (in.hasNextLine()) {
            out.add(Integer.parseInt(in.nextLine()));
        }
        return out;
    }

    /**
     * Add a delay to the code by a number of seconds
     *
     * @param sec the seconds to wait
     */
    private static void delay(int sec) {
        System.out.println("\r");
        try {
            Thread.sleep(1000 * sec);
        } catch (InterruptedException ex) {
        }
    }
}
