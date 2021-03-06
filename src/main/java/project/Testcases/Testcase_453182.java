package project.Testcases;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import project.excel.excelOperation;

import java.io.IOException;
import java.util.Map;

public class Testcase_453182 {
    private static final Logger log = LogManager.getLogger();
    public Testcase_453182(String sheetName, String rowNum) throws IOException, InvalidFormatException {
        log.info("Start of Execution : Testcase_453182");
        excelOperation excel = new excelOperation();
        try {
            Map<String, Map<String, String>> excelData = excel.getExcelData(sheetName);
            System.out.println(excelData.get(rowNum).get("textfield"));
            excel.setExcelData(sheetName,"TestCase Status",Integer.parseInt(rowNum),"PASS");
            System.out.println("Testcase_453182 Passed");
        } catch (Exception e) {
            log.error(e);
            excel.setExcelData(sheetName,"TestCase Status",Integer.parseInt(rowNum),"FAIL");
            System.out.println("Testcase_453182 Failed");

        }
    }
}