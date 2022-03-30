package review;

import GUI.MainGUI;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Spreadsheet class to handle lookups in a sheet
 *
 * @author Tyler Rose
 */
public class Spreadsheet {

    private static Spreadsheet sheet = null;
    private boolean isSetUp;

    private FileInputStream file;

    private XSSFWorkbook book;
    private HashMap<String, Integer> colTitles;
    private String fileLocation;

    /**
     * Get the instance of the Spreadsheet
     *
     * @return the spreadsheet instance
     */
    public static Spreadsheet getInstance() {
        if (sheet == null) {
            sheet = new Spreadsheet();
        }
        return sheet;
    }

    public static void resetInstance() {
        sheet = new Spreadsheet();
    }

    /**
     * Private default constructor for Singleton Spreadsheet object
     */
    private Spreadsheet() {
        isSetUp = false;
    }

    /**
     * Make sure the sheet vars have been set up before things can run
     */
    private void isSetUp() {
        if (!isSetUp) {
            throw new RuntimeException("Workbook Setup Error: Run the setup method first");
        }
    }

    /**
     * Set up the workbook variables
     *
     * @param fileLocation the location of the workbook to set up
     * @throws FileNotFoundException the workbook couldn't be located
     * @throws IOException the workbook couldn't be accessed
     */
    public void setupWorkbook(String fileLocation) throws FileNotFoundException, IOException {
        file = new FileInputStream(new File(fileLocation));
        book = new XSSFWorkbook(file);
        this.fileLocation = fileLocation;
    }

    /**
     * Set all variables for the spreadsheet including the column titles needed
     * for lookups and the CellWriter to edit cells
     *
     * @param workbookSetupThread the thread setting up the workbook
     * @param sheet the sheet number being set up
     * @throws InterruptedException The thread got interrupted
     */
    public void setupSpreadsheet(Thread workbookSetupThread, int sheet) throws InterruptedException {
        workbookSetupThread.join();
        isSetUp = true;
        MainGUI.println("Setting up Spreadsheet...");
        //get column titles
        colTitles = new HashMap<>();
        int i = 0;
        for (Cell cell : book.getSheetAt(sheet).getRow(0)) {
            colTitles.put((cell + "").toLowerCase(), i++);
        }
        CellWriter writer = CellWriter.getInstance();
        writer.setUpWriter(book, fileLocation);
        MainGUI.println("Finished Setting up Spreadsheet");
    }

    /**
     * Gets the sheet object by sheet number
     *
     * @param sheetNum the sheet number to get
     * @return the sheet at the given sheet numbers
     */
    public Sheet getSheet(int sheetNum) {
        isSetUp();
        return book.getSheetAt(sheetNum);
    }

    /**
     * Returns an ArrayList of all rows that have the same admin contact
     *
     * @param sheet the sheet to query
     * @param email the admin contact email to search for
     * @return the ArrayList of all rows with the given admin contact
     * @throws RuntimeException A null cell didn't get handled
     */
    public ArrayList<Row> getRowsByAdminContactEmail(Sheet sheet, String email) throws RuntimeException {
        isSetUp();
        //Get email column
        int col = colTitles.get(("Administrative Contact Email").toLowerCase());
        //Get all rows with the email that matches the provided one
        ArrayList<Row> rows = new ArrayList<>();
        int i = 0;
        for (Row row : sheet) {
            Cell cell = row.getCell(col);
            if (cell != null && getCellValue(cell).toLowerCase().equals(email.toLowerCase())) {
                rows.add(sheet.getRow(i));
            }
            i++;
            if (i >= 10000) {
                return rows;
            }
        }
        return rows;
    }

    /**
     * Get all rows with data in the column of the given title
     *
     * @param sheet the sheet to query
     * @param columnName the column name that must contain data
     * @return an ArrayList of rows that contain data in the column
     * @throws RuntimeException null cell error problems
     */
    public ArrayList<Row> getRowsByColumnName(Sheet sheet, String columnName) throws RuntimeException {
        isSetUp();
        //Get email column
        int col = colTitles.get(columnName.toLowerCase());
        //Get all rows with data in the column title that matches the provided one
        ArrayList<Row> rows = new ArrayList<>();
        int i = 0;
        for (Row row : sheet) {
            Spreadsheet mySpreadSheet = Spreadsheet.getInstance();
            String listingID = mySpreadSheet.getCellValue(mySpreadSheet.getCellByRowAndTitle(row, "Listing ID"));
            if (listingID.equals("")) {
                continue;
            }
            Cell cell = row.getCell(col);
            if (cell != null) {
                rows.add(sheet.getRow(i));
            }
            i++;
            if (i >= 10000) {
                return rows;
            }
        }
        return rows;
    }

    /**
     * Get a cell by its row and the column it is in
     *
     * @param row the row containing the cell
     * @param title the title of the column containing the cell
     * @return the cell at the given row and column
     * @throws java.lang.NullPointerException the cell at the requested column
     * didn't exist
     */
    public Cell getCellByRowAndTitle(Row row, String title) throws java.lang.NullPointerException {
        isSetUp();
        int columnIndex = colTitles.get(title.toLowerCase());
        try {
            return row.getCell(columnIndex);
        } catch (java.lang.NullPointerException npe) {
            throw new java.lang.NullPointerException("Column title not found:" + title);
        }
    }

    /**
     * Get the text in a cell
     *
     * @param cell the cell to read
     * @return the text read from the cell
     */
    public String getCellValue(Cell cell) {
        isSetUp();
        String value = "";
        FormulaEvaluator evaluator = book.getCreationHelper().createFormulaEvaluator();
        if (cell != null) {
            if (cell.toString().length() > 1 && cell.toString().contains(")")) {
                value = evaluator.evaluate(cell).getStringValue();
            } else {
                value = cell.toString();
            }
        }
        return value;
    }

    /**
     * Get the date
     *
     * @return the formatted current date
     */
    public static String getDateString() {
        SimpleDateFormat formatter = new SimpleDateFormat("M.dd.yy");
        Date date = new Date();
        //MainGUI.println(formatter.format(date));
        return formatter.format(date);
    }
}
