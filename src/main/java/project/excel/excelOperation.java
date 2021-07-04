package project.excel;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class excelOperation {

    private static final Logger log = LogManager.getLogger();
    public String currentDir = System.getProperty("user.dir");
    public String testReportPath = currentDir + File.separator + "TestReport" + File.separator + "data.xlsx";
    public Workbook wb = WorkbookFactory.create(new FileInputStream(testReportPath));

    public excelOperation() throws IOException, InvalidFormatException {
    }


    public String excelRead(String sheetName, int rowNum, int cellNum) throws IOException {
        String result = null;
        try {
            result = wb.getSheet(sheetName).getRow(rowNum).getCell(cellNum).getStringCellValue();
        } catch (NullPointerException e) {
            log.info("This is an empty cell");
            excelWrite(sheetName, rowNum, cellNum, " ");
        }
        return result;
    }


    public int excelReadInt(String sheetName, int rowNum, int cellNum) {
        double result = 0;
        try {
            result = wb.getSheet(sheetName).getRow(rowNum).getCell(cellNum).getNumericCellValue();
        } catch (NullPointerException e) {
            log.info("This is an empty cell");
        }
        return (int) result;
    }


    public void excelWrite(String sheetName, int rowNum, int cellNum, String data) throws IOException {
        wb.getSheet(sheetName).getRow(rowNum).createCell(cellNum).setCellValue(data);
        FileOutputStream fileWrite = new FileOutputStream(testReportPath);
        wb.write(fileWrite);
    }


    public int RowCount(String sheetName) {
        return wb.getSheet(sheetName).getLastRowNum();
    }


    public int CellCount(String sheetName, int rowNum) {
        return wb.getSheet(sheetName).getRow(rowNum).getLastCellNum();
    }


    public Map<String, Map<String, String>> getExcelDataAsMap(String sheetName) throws IOException {
        int rowCount = RowCount(sheetName);
        int cellCount = CellCount(sheetName, rowCount);
        Map<String, Map<String, String>> completeSheetData = new HashMap<String, Map<String, String>>();
        List<String> columnHeader = new ArrayList<String>();
        Sheet sheet = wb.getSheet(sheetName);
        Row row = sheet.getRow(0);
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            columnHeader.add(cellIterator.next().getStringCellValue());
        }
        for (int i = 1; i <= rowCount; i++) {
            Map<String, String> singleRowData = new HashMap<String, String>();
            Row HeaderRow = sheet.getRow(i);
            for (int j = 0; j < cellCount; j++) {
                Cell cell = HeaderRow.getCell(j);
                singleRowData.put(columnHeader.get(j), excelRead(sheetName, i, j));
            }
            completeSheetData.put(String.valueOf(i), singleRowData);
        }
        return completeSheetData;
    }


    public int GetCellNumber(String sheetName, String columnName) {
        int cellNumber = 0;
        Sheet sheet = wb.getSheet(sheetName);
        Row row = sheet.getRow(0);
        List<String> columnHeader = new ArrayList<String>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            columnHeader.add(cellIterator.next().getStringCellValue());
        }
        int rowCount = RowCount(sheetName);
        int cellCount = CellCount(sheetName, rowCount);
        for (int i = 1; i <= rowCount; i++) {
            for (int j = 0; j < cellCount; j++) {
                if (columnHeader.get(j).contains(columnName)) {
                    cellNumber = j;
                    break;
                }
            }
        }
        return cellNumber;
    }

}
