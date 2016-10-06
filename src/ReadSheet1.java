import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;


/**
 * Created by Jaye on 9/30/2016.
 */
public class ReadSheet1 implements Configuration{

    ExcelReader reader =  new ExcelReader();

    public HashMap finalOutput (String path) throws IOException {

        System.out.println("---------------------------- Readding sheet1 -----------------------------------------");
        boolean status = false;

        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(0);
        //Object [][] hardcodeValues = {{6,9,"Total"},{6,10,"Test Cut Requirement"}};


        HashMap <String,Object> sheet1  = new HashMap();
        try {
            reader.hardCodeValidator(sheet, sheet1HardcodeValues);
        } catch (RuntimeException ex){
            status = true;
        }

        try {
            sheet1.putAll(readTable1(path));
        } catch (RuntimeException ex){
            status = true;
        }

        try {
            sheet1.putAll(readTable2(path));
        } catch (RuntimeException ex){
            status = true;
        }

        System.out.println("Sheet 1 fully validated");
        return sheet1;
    }

    public HashMap readTable1 (String path) throws IOException {

        HashMap <String,Object> table1part1  = new HashMap();
        HashMap <String,Object> table1part2  = new HashMap();
        HashMap <String,Object> table1  = new HashMap();
        HashMap <String,Object> sizeMetrix  = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(0);

        /*
        Read Table 1 of the sheet 1
        */

        //String [] headerArray1 = {"Business Division","Style No","Season","Style Name","Silhouette"};
        //String [] headerArray2 = {"Customer Name/Category","Sample Type","Merchant Name","Garment Tech Name"};

        int [] rawHeaders = reader.rowHeaderValidator(sheet,1,1,sheet1Table1HeaderArray1);
        int [] rawheaders2 = reader.rowHeaderValidator(sheet,1,5,sheet1Table1HeaderArray2);

        table1part1 = reader.readRowHeaderTable(sheet, 1, rawHeaders,3,3 );
        table1part2 = reader.readRowHeaderTable(sheet,5,rawheaders2,8,8);
        table1.putAll(table1part1);
        table1.putAll(table1part2);

/** ----------------------------------------------------------------------------------------------------------------------- **/

        //String [] sizeMetrixColumns = {"XS","S","M","L","XL","XXL"};
        //String [] sizeMetrixHeaders = {"QTY","Aditional QTY"};

        int[] columnNumbers = reader.columnHeaderValidator(sheet,6,sheet1sizeMetrixColumns);
        int[] rowIndexes = reader.rowHeaderValidator(sheet,7,1,sheet1sizeMetrixHeaders);

        sizeMetrix =reader.readColumnAndRowHeaderTable(sheet,6, 1, columnNumbers, 3, 8, rowIndexes);
        table1.put("sizeMetrix", sizeMetrix);

        table1.put(reader.getValue(sheet.getRow(6).getCell(9)).toString(),reader.getValue(sheet.getRow(7).getCell(9)));
        table1.put(reader.getValue(sheet.getRow(6).getCell(10)).toString(),reader.getValue(sheet.getRow(7).getCell(10)));
        return table1;


    }

    public HashMap readTable2 (String path) throws IOException {



        HashMap <String,Object> part1  = new HashMap();
        HashMap <String,Object> part2  = new HashMap();
        HashMap <String,Object> part3  = new HashMap();
        HashMap <String,Object> table2  = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(0);

        //String [][] headers = {{"ITEM"},{"IM"},{"Sub IM"},{"ITEM COLOR"},{"Sub color"},{"CW"},{"WIDTH/SIZE"},{"ITEM QTY"},{"Fabric Face Side"}, {"USAGE"},{"SUPPLIER"},{"DESCREPTION"},{"COMMENTS AND REMARKS"}};

        int [][] headerIndexes = reader.validator(sheet,10,sheet1table2Headers);


        //int [][] a = {{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}};
        part1 = reader.readComplexColumnHeaderTable(sheet, 10, headerIndexes, 11,16);
        part2 = reader.readComplexColumnHeaderTable(sheet, 10, headerIndexes, 17,23);
        part3 = reader.readComplexColumnHeaderTable(sheet, 10, headerIndexes, 24,30);

        table2.put(reader.getValue(sheet.getRow(11).getCell(1)).toString(), part1);
        table2.put(reader.getValue(sheet.getRow(17).getCell(1)).toString(), part2);
        table2.put(reader.getValue(sheet.getRow(24).getCell(1)).toString(),part3);

        return table2;
    }

}
