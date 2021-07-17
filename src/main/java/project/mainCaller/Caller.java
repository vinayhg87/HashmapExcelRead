package project.mainCaller;

import project.excel.excelOperation;
import java.lang.reflect.Constructor;
import java.util.Map;

public class Caller {

    public static void main(String[] args) throws Exception {

        excelOperation excel = new excelOperation();
        /* Life Events Test cases */
        String sheetName = "Life Events";
        Map<String, Map<String, String>> excelData = excel.getExcelData(sheetName);
        //System.out.println(excelData);
        //int rowCount = excel.RowCount(sheetName);
        for (String rowNum : excelData.keySet()) {
            if (!(excelData.get(rowNum).get("TestCase Status")).equalsIgnoreCase("PASS")
                    || (excelData.get(rowNum).get("TestCase Status")) == null) {

                String ClassName = "project.Testcases." + excelData.get(rowNum).get("TestCase Number");
                Class<?> FormClass = Class.forName(ClassName);
                Constructor<?> constructor = FormClass.getConstructor(String.class, String.class);
                constructor.newInstance(sheetName, rowNum);//caller
            } else {
                System.out.println("The test case " + excelData.get(rowNum).get("TestCase Number")
                                                                    + " was executed with status PASS.");
            }
        }
    }
}

