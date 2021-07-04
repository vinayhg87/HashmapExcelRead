package project.Testcases;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import project.excel.excelOperation;
import java.util.Map;

public class Testcase_453178 {
    private static final Logger log = LogManager.getLogger();
    public Testcase_453178(String sheetname, String rowNum) throws Exception {
        log.info("Start of Execution : Testcase_453178");
        excelOperation excel = new excelOperation();
        try {
            Map<String, Map<String, String>> excelData = excel.getExcelAsMap(sheetname);
            System.out.println(excelData.get(rowNum).get("LastName1"));
            excel.excelWrite(sheetname, Integer.parseInt(rowNum),
                    excel.GetCellNumber(sheetname, "TestCase Status"), "PASS");
            System.out.println("Testcase_453178 Passed");
        } catch (Exception e) {
            log.error(e);
            excel.excelWrite(sheetname, Integer.parseInt(rowNum),
                    excel.GetCellNumber(sheetname, "TestCase Status"), "FAIL");
            System.out.println("Testcase_453178 Failed");
        }
    }
}
