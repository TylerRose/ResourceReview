package review;

import GUI.LoginGUI;
import GUI.MainGUI;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.util.ArrayList;
import java.util.Scanner;

/**
 * Main file of RR that handles the program flow, major steps, and error
 * handling
 *
 * @author Tyler Rose
 */
public class Source {

    public static boolean close = false;

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
    public static void SourceMain(String[] args) {
        //SendEmail.sendAnEmail(new Email("", "", "")); //handle the arguments and enable test/error modes
        if (args.length == 0) {
            //set default argument values
            SendEmail.testMode = true;
        } else {
            //Set test mode (send all internal) if first argument is test= true/false
            switch (args[0].toLowerCase()) {
                case "test=readonly":
                    RRMain.writer.readOnly = true;
                case "test=true":
                    SendEmail.testMode = true;
                    break;
                case "test=false":
                    //if test mode is off, run a confirmation input step
                    SendEmail.testMode = false;
                    do {
                        MainGUI.print("Running and sending to RR contacts.\n**If re-sending errors** the listing IDs will be emailed using the data of the month/year you are about to select.\nType 'yes' to confirm production run: ");
                    } while (!in.nextLine().toLowerCase().contains("yes"));
                    MainGUI.println("");
                    break;
                default:
                    MainGUI.println("Unknown arguments, running in TEST MODE");
                    SendEmail.testMode = true;
                    break;
            }
            if (args.length > 1) {
                switch (args[1].toLowerCase()) {
                    case "errors":
                        RRMain.errorsOnly = true;
                        break;
                }
            }
        }
        //MainGUI.getInstance().startGUI();
        //Wait to return until GUI closes
        while (!close) {
            delay(0.5);
        }
    }

    public static void startReview() {
        getInput();
        if (!RRMain.ReviewSteps(RRMain.errorsOnly)) {
            MainGUI.println("An error has occured: see details above.");
            MainGUI.print("The sheet has not been modified, no reviews have been run.");
            MainGUI.println("\n\n");
        }
        //reset all variables to defaults to run again
        resetVars();
    }

    public static void resetVars() {
        LoginGUI.retrying = false;

        RRMain.doneIDs = new ArrayList<>();

        SendEmail.index = 0;
        SendEmail.error = 0;
        SendEmail.username = "";
        SendEmail.sentCount = 0;
        SendEmail.session = null;
        SendEmail.setPassword("");

        //reset instances
        EmailManager.resetInstance();
        ErrorTracker.resetInstance();
        Spreadsheet.resetInstance();
    }

    /**
     * Get year, month, and specialist information to know what is being run and
     * where it is located
     *
     * @return the sheet number that was provided
     */
    private static void getInput() {

        MainGUI gui = MainGUI.getInstance();
        //MainGUI.println("Please enter the year: ");
        //year = in.nextInt();
        if (gui.getTxtYear().equals("Year")) {
            RRMain.year = -1;
        } else {
            RRMain.year = Integer.parseInt(gui.getTxtYear());
        }
        RRMain.excelPath = RRMain.resourceReviewsPath + "Excel\\" + (RRMain.year == -1 ? "0test0" : RRMain.year) + "\\";
        //set up the sheet in a separate thread to load data while getting input
        RRMain.workbookSetup = new Thread() {
            @Override
            public void run() {
                try {
                    RRMain.mySpreadSheet.setupWorkbook(RRMain.excelPath + "\\" + new File(RRMain.excelPath).list()[0]);
                } catch (IOException ex) {
                    Source.printError("The workbook could not be set up, the file was inaccessable.");
                }
            }
        };
        RRMain.workbookSetup.start();

        //Create the path if it doesn't exist
        new File(RRMain.excelPath).mkdirs();
        //Get month and specialist details
        //MainGUI.println("Please enter the Month number: ");
        //sheetNo = in.nextInt() - 1;
        RRMain.sheetNo = Integer.parseInt(gui.getTxtMonth()) - 1;
//        MainGUI.print("Please select the specialist:");
//        int i = 1;
//        for (String s : specialistList) {
//            MainGUI.print(" (" + i++ + ") " + s);
//        }
//        MainGUI.println("");
//        int selected = in.nextInt();
        //String name = specialistList.get(selected - 1);
        String name = gui.getSpnSpecialist();
        RRMain.specialistInitials = name.substring(0, 1).concat(name.substring(name.indexOf(" ") + 1, name.indexOf(" ") + 2));
    }

    /**
     * A method to print formatted errors to console
     *
     * @param message the message to print
     */
    public static void printError(String message) {

        MainGUI.println("\n******");
        MainGUI.println("ERROR: " + message);
        MainGUI.println("******\n");
    }

    /**
     * Add a delay to the code by a number of seconds
     *
     * @param sec the seconds to wait
     */
    public static void delay(double sec) {
        //MainGUI.println("\r");
        try {
            Thread.sleep((long) (1000 * sec));
        } catch (InterruptedException ex) {
        }
    }
}
