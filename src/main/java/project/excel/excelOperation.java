package project.excel;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
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

    public String excelRead(String sheetName, int rowNum, int cellNum) throws Exception {
        String result = null;
        try {
            Cell cellType = wb.getSheet(sheetName).getRow(rowNum).getCell(cellNum);
            if (cellType.getCellType() == cellType.CELL_TYPE_NUMERIC) {
                result = NumberToTextConverter.toText(cellType.getNumericCellValue());
            } else if (cellType.getCellType() == cellType.CELL_TYPE_STRING) {
                result = cellType.getStringCellValue();
            } else if (cellType.getCellType() == cellType.CELL_TYPE_BLANK) {
                System.out.println("This is an Blank cell");
            }
        } catch (NullPointerException e) {
            System.out.println("This is an empty cell");
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

    public int sheetCount() {
        return wb.getNumberOfSheets();
    }

    public String sheetName(int sheetIndex) {
        return wb.getSheetName(sheetIndex);
    }


    public int CellCount(String sheetName, int rowNum) {
        return wb.getSheet(sheetName).getRow(rowNum).getLastCellNum();
    }


    public Map<String, Map<String, String>> getExcelData(String sheetName) throws Exception {
        int lastrowCount = RowCount(sheetName);
        int lastcellCount = CellCount(sheetName, lastrowCount);
        Sheet sheet = wb.getSheet(sheetName);
        Map<String, Map<String, String>> fullSheetData = new HashMap<String, Map<String, String>>();
        List<String> columnHeader = new ArrayList<String>();
        Row row = sheet.getRow(0);
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            columnHeader.add(cellIterator.next().getStringCellValue());
        }
        for (int i = 1; i <= lastrowCount; i++) {
            Map<String, String> singleRowData = new HashMap<String, String>();
            Row HeaderRow = sheet.getRow(i);
            for (int j = 0; j < lastcellCount; j++) {
                Cell cell = HeaderRow.getCell(j);
                singleRowData.put(columnHeader.get(j), excelRead(sheetName, i, j));
            }
            fullSheetData.put(String.valueOf(i), singleRowData);
        }
        return fullSheetData;
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


    public void setExcelData(String sheetName, String columnName, int RowNum, String data) throws IOException {
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
        excelWrite(sheetName, RowNum, cellNumber, data);
    }


    public boolean fileCheck() {
        boolean fileStatusCheck = false;
        File fileExists = new File(testReportPath);
        boolean fileLoc = fileExists.exists();

        if (fileLoc) {
            try {
                FileOutputStream checkFile = new FileOutputStream(testReportPath, true);
                checkFile.close();
                fileStatusCheck = true;
            }

            catch (Exception e) {
                log.fatal("FATAL ERROR : " + "data.xlsx" + " is already open. "
                        + "Please close it and run the program again.");
                fileStatusCheck = false;
            }
        } else {
            log.fatal("FATAL ERROR : " + "data.xlsx" + " file not found. " + "File should be at " + testReportPath);
            System.exit(1);
        }
        return fileStatusCheck;
    }
}
