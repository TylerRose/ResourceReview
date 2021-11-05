
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A class that writes text to a Cell
 *
 * @author Tyler Rose
 */
public class CellWriter {

    private static CellWriter writer;
    private static XSSFWorkbook book;
    private static String fileLocation;
    private static boolean isSetUp;

    /**
     * Private constructor for Singleton CellWriter
     */
    private CellWriter() {
        isSetUp = false;
    }

    /**
     * Get the CellWriter instance
     *
     * @return an instance of CellWriter
     */
    public static CellWriter getInstance() {
        if (writer == null) {
            writer = new CellWriter();
            return writer;
        } else {
            return writer;
        }
    }

    /**
     * Write changes to the book and close the CellWriter object. Must be
     * re-initialized before it can be used again.
     *
     * @throws IOException Couldn't save/close the book. Perhaps it was open.
     */
    public void closeWriter() throws IOException {
        isSetUp();
        //initialize out as the file output
        try (FileOutputStream out = new FileOutputStream(fileLocation, false)) {            
            //write and close the book through file stream out
            book.write(out);
        }
        writer = null;
        book = null;
        fileLocation = null;
        isSetUp = false;
    }

    /**
     * Set up the CellWriter with a workbook and save file location
     *
     * @param book the workbook to write to
     * @param fileLocation the location of the workbook
     */
    public void setUpWriter(XSSFWorkbook book, String fileLocation) {
        isSetUp = true;
        CellWriter.book = book;
        CellWriter.fileLocation = fileLocation;
    }

    /**
     * Check if the cell writer was set up
     */
    private void isSetUp() {
        if (!isSetUp) {
            throw new RuntimeException("The Cell Writer wasn't set up");
        }
    }

    /**
     * Set the text of a cell
     *
     * @param cell the cell to exit
     * @param newText the new text of the cell
     * @return the cell with the new text
     * @throws IllegalArgumentException The provided cell to edit was null
     */
    public Cell setCellText(Cell cell, String newText) throws IllegalArgumentException {
        isSetUp();
        if (cell == null) {
            throw new IllegalArgumentException();
        }
        cell.setCellValue(newText.trim());
        return cell;
    }

    /**
     * Append text to the end of a cell
     *
     * @param cell the cell to exit
     * @param addedText the text to add to the cell
     * @return the cell with the new text
     * @throws IllegalArgumentException The provided cell to edit was null
     */
    public Cell appendCellText(Cell cell, String addedText) throws IllegalArgumentException {
        isSetUp();
        if (cell == null) {
            throw new IllegalArgumentException();
        }
        String appended = getCellValue(cell) + addedText;
        setCellText(cell, appended.trim());
        return cell;
    }

    /**
     * Get the text of a cell or evaluate it's formula and get the result
     *
     * @param cell the cell read
     * @return the text that was read or the result of the cell's formula
     * @throws RuntimeException The provided cell was null
     */
    private String getCellValue(Cell cell) throws RuntimeException {
        isSetUp();
        String value = "";
        FormulaEvaluator evaluator = book.getCreationHelper().createFormulaEvaluator();
        if (cell.toString().length() > 1 && cell.toString().contains(")")) {
            value = evaluator.evaluate(cell).getStringValue();
        } else {
            value = cell.toString();
        }
        return value;
    }
    
    /**
     * Evaluate the formula in a cell. The formula will remain in the cell and
     * Excel will display the resulting value.
     * 
     * @param cell the cell containing the formula to evaluate
     */
    public void refreshCell(Cell cell){
        isSetUp();
        FormulaEvaluator evaluator = book.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateFormulaCell(cell);
    }
}
