
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author trose
 */
public class CellWriter {

    private static CellWriter writer;
    private static XSSFWorkbook book;
    private static String fileLocation;
    private static boolean isSetUp;

    private CellWriter() {
        isSetUp = false;
    }

    public static CellWriter getInstance() {
        if (writer == null) {
            writer = new CellWriter();
            return writer;
        } else {
            return writer;
        }
    }

    public void closeWriter() throws IOException {
        isSetUp();
        try (FileOutputStream out = new FileOutputStream(fileLocation, false)) {
            book.write(out);
        }
        writer = null;
        book = null;
        fileLocation = null;
        isSetUp = false;
    }

    public void setUpWriter(XSSFWorkbook book, String fileLocation) {
        isSetUp = true;
        CellWriter.book = book;
        CellWriter.fileLocation = fileLocation;
    }

    private void isSetUp() {
        if (!isSetUp) {
            throw new RuntimeException("The Cell Writer couldn't wasn't set up");
        }
    }

    public Cell setCellText(Cell cell, String newText) throws IllegalArgumentException {
        isSetUp();
        if (cell == null) {
            throw new IllegalArgumentException();
        }
        cell.setCellValue(newText.trim());
        return cell;
    }

    public Cell appendCellText(Cell cell, String addedText) throws IllegalArgumentException {
        isSetUp();
        if (cell == null) {
            throw new IllegalArgumentException();
        }
        String appended = getCellValue(cell) + addedText;
        setCellText(cell, appended.trim());
        return cell;
    }

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

}
