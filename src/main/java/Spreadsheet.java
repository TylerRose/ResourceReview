
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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author trose
 */
public class Spreadsheet {

    private static Spreadsheet sheet = null;
    private boolean isSetUp = false;

    private FileInputStream file;

    private XSSFWorkbook book;
    private HashMap<String, Integer> colTitles;
    private String fileLocation;

    public static Spreadsheet getInstance() {
        if (sheet == null) {
            sheet = new Spreadsheet();
        }
        return sheet;
    }

    private Spreadsheet() {

    }

    private void isSetUp() {
        if (!isSetUp) {
            throw new RuntimeException("Run the setup method first");
        }
    }

    public void setupWorkbook(String fileLocation) throws FileNotFoundException, IOException {
        file = new FileInputStream(new File(fileLocation));
        book = new XSSFWorkbook(file);
        this.fileLocation = fileLocation;
    }

    public void setupSpreadsheet(Thread workbookSetupThread, int sheet) throws InterruptedException {
        workbookSetupThread.join();
        isSetUp = true;
        System.out.println("Setting up Spreadsheet...");
        //get column titles
        colTitles = new HashMap<>();
        int i = 0;
        for (Cell cell : book.getSheetAt(sheet).getRow(0)) {
            colTitles.put(cell + "", i++);
        }
        CellWriter writer = CellWriter.getInstance();
        writer.setUpWriter(book, fileLocation);
        System.out.println("Finished Setting up Spreadsheet");
    }

    public Sheet getSheet(int sheetNum) {
        isSetUp();
        return book.getSheetAt(sheetNum);
    }

    public ArrayList<Row> getRowsByAdminContactEmail(Sheet sheet, String email) throws RuntimeException {
        isSetUp();
        //Get email column
        int col = colTitles.get("Administrative Contact Email");
        //Get all rows with the email that matches the provided one
        ArrayList<Row> rows = new ArrayList<>();
        int i = 0;
        for (Row row : sheet) {
            Cell cell = row.getCell(col);
            if (cell != null && getCellValue(cell).equals(email)) {
                rows.add(sheet.getRow(i));
            }
            i++;
            if (i >= 10000) {
                return rows;
            }
        }
        return rows;
    }

    public ArrayList<Row> getRowsByColumnName(Sheet sheet, String columnName) throws RuntimeException {
        isSetUp();
        //Get email column
        int col = colTitles.get(columnName);
        //Get all rows with the column title that matches the provided one
        ArrayList<Row> rows = new ArrayList<>();
        int i = 0;
        for (Row row : sheet) {
            Cell cell = row.getCell(col);
            if (cell != null) {
                rows.add(sheet.getRow(i));
            }
            i++;
            if(i>= 10000){
                return rows;
            }
        }
        return rows;
    }

    public Cell getCellByRowAndTitle(Row row, String title) throws java.lang.NullPointerException {
        isSetUp();
        //System.out.println(title);
        int columnIndex = colTitles.get(title);
        //System.out.println("Getting title: " + title + " in row " + row.getRowNum());
        try {
            //System.out.println("Trying getCell on row " + row.getRowNum() + " column " + columnIndex + " (" + title + ") and returning: " + row.getCell(columnIndex));
            return row.getCell(columnIndex);
        } catch (java.lang.NullPointerException npe) {
            //System.out.println("Column title gave null:" + title);
            throw new java.lang.NullPointerException("Column title gave null:" + title);
        }
    }

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

    public static String getDateString() {
        SimpleDateFormat formatter = new SimpleDateFormat("M.dd.yy");
        Date date = new Date();
        //System.out.println(formatter.format(date));
        return formatter.format(date);
    }
}
