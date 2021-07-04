package project.Testcases;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import project.excel.excelOperation;
import java.io.IOException;
import java.util.Map;

public class Testcase_453181 {
    private static final Logger log = LogManager.getLogger();
    public Testcase_453181(String sheetName, String rowNum) throws IOException, InvalidFormatException {
        log.info("Start of Execution : Testcase_453181");
        excelOperation excel = new excelOperation();
        try {
            Map<String, Map<String, String>> excelData = excel.getExcelAsMap(sheetName);
            System.out.println(excelData.get(rowNum).get("LastName1"));
            excel.excelWrite(sheetName, Integer.parseInt(rowNum),
                    excel.GetCellNumber(sheetName, "TestCase Status"), "PASS");
            System.out.println("Testcase_453181 Passed");
        } catch (Exception e) {
            log.error(e);
            excel.excelWrite(sheetName, Integer.parseInt(rowNum),
                    excel.GetCellNumber(sheetName, "TestCase Status"), "FAIL");
            System.out.println("Testcase_453181 Failed");
        }
    }
}
