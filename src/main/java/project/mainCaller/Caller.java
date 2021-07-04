package project.mainCaller;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import project.excel.excelOperation;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.Map;

public class Caller {

    public static void main(String[] args) throws IOException, InvalidFormatException,
            ClassNotFoundException, NoSuchMethodException, InvocationTargetException,
            InstantiationException, IllegalAccessException {

        excelOperation excel = new excelOperation();

        /* Life Events Test cases */
        String sheetName = "Life Events";
        Map<String, Map<String, String>> excelData = excel.getExcelAsMap(sheetName);
        int rowCount = excel.RowCount(sheetName);
        for (String rowNum : excelData.keySet()) {
            if (!(excelData.get(rowNum).get("TestCase Status")).equalsIgnoreCase("PASS")
                    || (excelData.get(rowNum).get("TestCase Status")) == null) {

                String ClassName = "project.Testcases." + excelData.get(rowNum).get("TestCase Number");
                Class<?> FormClass = Class.forName(ClassName);
                Constructor<?> constructor = FormClass.getConstructor(String.class, String.class);
                constructor.newInstance(sheetName, rowNum);//caller
            } else {
                System.out.println("The test case " + excelData.get(rowNum).get("TestCase Number") + " was executed with status PASS.");
            }
        }
    }
}
