
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.*;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author trose
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

    private static void delayDot(int dots) {
        System.out.println("\r");
        for (int i = 0; i < dots; i++) {
            try {
                Thread.sleep(1000);
            } catch (InterruptedException ex) {
                //Just dot 
            }

            //System.out.print('.');
            //System.out.flush();
        }
        //System.out.println("\r\n");
    }
    ArrayList<String> doneIDs = new ArrayList<>();

    private static String fileLocation;// = "C:\\Excel\\Test Excel Sheet-2021.xlsx";
    //private static FileInputStream file;
    //private static Workbook workbook;
    private static final Scanner in = new Scanner(System.in);

    public static void main(String[] args) {

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
                    SendEmail.testMode = false;
                    System.out.println("Running and sending to RR contacts. Type 'yes' to confirm production run: ");
                    while (!in.nextLine().toLowerCase().contains("yes")) {
                        System.out.print("Running and sending to RR contacts. Type 'yes' to confirm production run: ");
                    }
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

        //Load settings
        try {
            loadSettings();
        } catch (RuntimeException e) {
            System.out.println("ERROR: " + e.getMessage());
        }
//        excelPath = resourceReviewsPath + "Excel\\" + 2021 + "\\";
//        File docPath = new File(excelPath);
//        fileLocation = excelPath + "\\" + docPath.list()[0];
//        Spreadsheet mysheet = new Spreadsheet(fileLocation);
        int sheetNo = getInput();
        //Give the month to the powershell file to use for confirDmation email
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
            System.out.println("ERROR: Missing permissions to write to file! (Path: " + powershellScript + "\\files\\" + monthNames[sheetNo] + ")");
        }
        System.out.print("\nProcessing");
        delayDot(3);

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

        Sheet sheet = null;
        if (errorsOnly) {
            //Begin re-running only errored emails
            RunErrorsOnly();
            //End program when errors are finished
        } else {
            try {
                //Beginnin running normal execution
                sheet = RunResourceReview(sheetNo);
            } catch (IllegalArgumentException ex) {
                System.out.println("ERROR: The tabs were out of order! Please re-order the tabs and run again.");
            } catch (FileNotFoundException ex) {
                System.out.println("ERROR: The spreadsheet could not be set up, the file was inaccessable.");
            }
        }

        System.out.println("Finished processing");

        System.out.print("\nSending emails");
        delayDot(3);
        try {
            sendEmails();
        } catch (IOException ex) {
            System.out.println("ERROR: Unable to start powershell processes. Please run manually in the PS folder.");
        }

        ErrorTracker errors = reportErrors();

        if (sheet != null) {
            System.out.print("\nUpdating dates for successfull entries");
            delayDot(3);
            try {
                updateDates(sheet, errors);
            } catch (RuntimeException ex) {
                System.out.println("A runtime exception occured:\n" + ex.getMessage());
                ex.printStackTrace();
            }
            try {
                writer.closeWriter();
                System.out.println("Finished updating dates.");
            } catch (IOException ex) {
                System.out.println("ERROR: Unable to save and close the sheet. Dates have not been updated. Make sure the sheet is closed before running.");
            }

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
            System.out.println("The sheet to run wasn't identified. Verify all tab names and file locations.");
            System.out.println("Nothing has been run.");
        }
    }

    private static Sheet RunResourceReview(int sheetNo) throws IllegalArgumentException, FileNotFoundException, IllegalArgumentException {
        ParseEmailFormat parse = new ParseEmailFormat(mySpreadSheet.getSheet(sheetNo), resourceReviewsPath);
        //Parse through the email addresses, combining all agencies per address before moving to the next
        Sheet sheet = mySpreadSheet.getSheet(sheetNo);
        //Make sure tabs are in the correct order before running the sheet
        CheckTabOrder(sheet, sheetNo);
        String prevEmail = "----";
        String currEmail = "";
        ArrayList<String> done = new ArrayList<>();
        int maxRow = 10000;
        int curRow = 0;
        for (Row row : sheet) {
            // System.out.println("Row: " + row.getRowNum());
            try {
                Cell cell = mySpreadSheet.getCellByRowAndTitle(row, "Agency ID");
                Cell completed = mySpreadSheet.getCellByRowAndTitle(row, "Complete");
                if (cell != null && mySpreadSheet.getCellValue(completed).equals("")) {
//                    System.out.println("Completed: " + mySpreadSheet.getCellValue(completed));
//                    System.out.println("Prev: " + prevEmail);
//                    System.out.println("Curr: " + currEmail);
//                    System.out.println("Done: " + done.toString());
//                    System.out.println("\n\n");
                    currEmail = (mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Administrative Contact Email"))).toLowerCase();
                    if (currEmail.length() > 3 && !currEmail.equals(prevEmail) && !done.contains(currEmail)) {
                        done.add(currEmail);
                        parse.parseRowsByEmail(mySpreadSheet.getSheet(sheetNo), currEmail, specialistInitials);
                        prevEmail = currEmail;
                    }
                }
            } catch (IOException | RuntimeException e) {
                //System.out.println(e);
                //System.out.println(Arrays.toString(e.getStackTrace()));
            }
            curRow++;
            if (curRow >= maxRow) {
                System.out.println("Hit 10k rows...");
                return sheet;
            }
        }
        return sheet;
    }

    private static void CheckTabOrder(Sheet sheet, int sheetNo) throws IllegalArgumentException {
        String fullTabName = sheet.getSheetName().toLowerCase().replace(" ", "").replace("_", "").replace("-", "");
        String lookingFor = "tab" + (sheetNo + 1);
        char nextChar;
        if (fullTabName.length() > fullTabName.indexOf(lookingFor) + lookingFor.length() + 1) {
            nextChar = fullTabName.charAt(fullTabName.indexOf(lookingFor) + lookingFor.length() + 1);
        } else {
            nextChar = 'a';
        }
        if (!(fullTabName).contains(lookingFor) || Character.isDigit(nextChar)) {
            System.out.println("Tabs our of order!");
            throw new IllegalArgumentException("Tabs out of order!");
        }
    }

    private static int getInput() {
        int sheetNo;
        System.out.println("Please enter the year: ");
        year = in.nextInt();
        excelPath = resourceReviewsPath + "Excel\\" + year + "\\";
        workbookSetup = new Thread() {
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

    private static void updateDates(Sheet sheet, ErrorTracker errors) throws RuntimeException {
        Cell currCell;
        String currCellText;
        int columnNum = mySpreadSheet.getCellByRowAndTitle(sheet.getRow(0), "Dates Contacted").getColumnIndex();
        int maxRow = 10000;
        int curRow = 0;
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
    //    //Old Send emails using the output folder with individual files
    //    private static void sendEmails() throws IOException {
    //        File scriptLoc = new File(resourceReviewsPath + "ps\\");
    //        File outputFolder = new File(resourceReviewsPath + "Outputs");
    //        //System.out.println(outputFolder.listFiles());
    //        if (outputFolder.listFiles().length > 0) {
    //            File current;
    //            Scanner in;
    //            String to;
    //            String email;
    //            String subject;
    //            for (int k = 0; k < outputFolder.listFiles().length - 1; k++) {
    //                email = "";
    //                current = outputFolder.listFiles()[k];
    //                in = new Scanner(new FileInputStream(new File(current.getAbsolutePath())));
    //                to = in.nextLine().substring("To- ".length());
    //                subject = in.nextLine().substring("Subject- ".length());
    //                while (in.hasNextLine()) {
    //                    email += in.nextLine();
    //                }
    //                SendEmail.addEmail(scriptLoc, to, "resourcereviews@homage.org", subject, email);
    //            }
    //            SendEmail.sendAll();
    //        }
    //    }
    //New sendEmails using the email manager

    private static void sendEmails() throws IOException {
        File scriptLoc = new File(resourceReviewsPath + "ps\\");
        EmailManager eManager = EmailManager.getInstance();
        ArrayList<Email> emailList = eManager.getEmails();
        for (Email e : emailList) {
            SendEmail.addEmail(scriptLoc, "resourcereviews@homage.org", e);
        }
        SendEmail.sendAll();
    }

    //Loads the settings from the setting file
    private static void loadSettings() {
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

    private static void RunErrorsOnly() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
}
