package test;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * This class executes read, write data to excel file
 */

public class XlsHandler {
    private List<List<String>> xlsSheet;
    private File sourceFile;
    private File outputFile;

    public XlsHandler(List<List<String>> xlsSheet, File sourceFile) {
        this.xlsSheet = xlsSheet;
        this.sourceFile = sourceFile;
        this.outputFile = new File(sourceFile.getParent() + File.separator + "outputFile.xls");
    }

    public void readXlsSourceFile() {
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(this.sourceFile));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;
            int rows = 11; // No of rows
//            rows = sheet.getPhysicalNumberOfRows();
            int cols = sheet.getRow(0).getPhysicalNumberOfCells(); // No of columns
            for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    this.xlsSheet.add(new ArrayList<>());
                } else {
                    continue;
                }
                List<String> currentRaw = this.xlsSheet.get(r);
                if (row != null) {
                    for (int c = 0; c < cols; c++) {
                        cell = row.getCell(c);
                        currentRaw.add(cell.toString());
                    }
                }
            }
        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    public void printXlsSourceFile() {
        for (List<String> row : this.xlsSheet) {
            System.out.println(row.get(2) + " " + row.get(4) + " " + row.get(16) + " " + row.get(17) + " ");
        }
    }

    public void writeXlsOutputFile() {
        Cell cell;
        Row row;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Resulted sheet");
            HSSFFont font = workbook.createFont();
            font.setBold(true);
            font.setItalic(true);
            font.setFontHeight((short) 240);
            HSSFCellStyle style = workbook.createCellStyle();
            style.setFont(font);
            for (int rowIndex = 0; rowIndex < this.xlsSheet.size(); rowIndex++) {
                row = sheet.createRow(rowIndex);
                List<String> currentRow = this.xlsSheet.get(rowIndex);
                for (int cellIndex = 0; cellIndex < currentRow.size(); cellIndex++) {
                    cell = row.createCell(cellIndex, CellType.STRING);
                    cell.setCellValue(currentRow.get(cellIndex));
                    if (rowIndex == 0) cell.setCellStyle(style);
                }
            }
            workbook.write(new FileOutputStream(this.outputFile));
        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    public static void main(String[] args) {
        File sourceFile = new File("src/test/new.xls");
        XlsHandler x = new XlsHandler(new ArrayList<>(), sourceFile);
        x.readXlsSourceFile();
        x.writeXlsOutputFile();
    }
}
