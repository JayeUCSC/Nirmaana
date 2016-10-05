import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

public class Main {

    public static void main(String[] args) throws IOException {

/*
        System.out.println("Hello World!");

        File excelFile = new File("F:\\Intern - nCinga\\Dev\\2nd one\\JC\\output Jsons\\865834  RM Chart 2nd Proto.xlsx");
        FileInputStream inputFile = new FileInputStream(excelFile);
        XSSFWorkbook workBook = new XSSFWorkbook(inputFile);
        XSSFSheet workSheet = workBook.getSheetAt(3);

        int[] rowIndexes = {2,3,4,5,6,7,8,9,10,11,12,13};

        //HashMap data = new ExcelReader().readColumnAndRowHeaderTable(workSheet,7, 0, columnNumbers, 8, 9, rowIndexes);

        HashMap data = new ExcelReader().readColumnHeaderTable(workSheet, 10, rowIndexes, 11, 16,11);
        new JsonWriter().write(data, "C:\\Users\\Jaye\\Desktop\\Test.json");

*/
      String path = "F:\\Intern - nCinga\\Dev\\2nd one\\JC\\865834  RM Chart 2nd Proto.xlsx";

       HashMap sheet1 = new ReadSheet1().finalOutput(path);
       new JsonWriter().write(sheet1, "C:\\Users\\Jaye\\Desktop\\Sheet1.json");

        HashMap sheet2 = new ReadSheet2().finalOutput(path);
        new JsonWriter().write(sheet2, "C:\\Users\\Jaye\\Desktop\\Sheet2.json");

        HashMap sheet3 = new ReadSheet3().finalOutput(path);
        new JsonWriter().write(sheet3, "C:\\Users\\Jaye\\Desktop\\Sheet3.json");

        HashMap sheet4 = new ReadSheet4().finalOutput(path);
        new JsonWriter().write(sheet4, "C:\\Users\\Jaye\\Desktop\\Sheet4.json");

    }
}
