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
            Map<String, Map<String, String>> excelData = excel.getExcelData(sheetname);
            System.out.println(excelData.get(rowNum).get("FirstName2"));
            excel.setExcelData(sheetname,"TestCase Status",Integer.parseInt(rowNum),"PASS");
            System.out.println("Testcase_453178 Passed");
        } catch (Exception e) {
            log.error(e);
            excel.setExcelData(sheetname,"TestCase Status",Integer.parseInt(rowNum),"FAIL");
            System.out.println("Testcase_453178 Failed");
        }
    }
}
