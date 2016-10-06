/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author acer
 */
public class ReadSheet2 implements Configuration{


    ExcelReader reader =  new ExcelReader();

    public HashMap finalOutput (String path) throws IOException {

        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(2);
  /*      Object [][] hardcodeValues = {{2,5,"SAMPLE TYPE"},{6,5,"EARLY COMPETION TIME"},{3,13,"Add. Cuts"}, {3,14,"Total"},{6,14,"SAMPLE PLAN END DATE & TIME"},
                {14,0,"Confirmation From Stores"},{14,10,"DATE",14,15,"TIME:"},{15,0,"Verification (G.Tech)"},{15,10,"DATE"},{15,15,"TIME:"},
                {17,0,"CUTTING TEAM :"},{17,15,"CUT DATE :"},{18,0,"CUTTER/S NAME :"},{18,15,"CUT TIME :"} ,
                {26,0,"CPI DONE BY :"},{26,16,"VERIFICATION DONE BY :"},{27,0,"COMMENTS :"},{27,16,"COMMENTS :"}};
*/
        System.out.println("---------------------------- Reading sheet2 -----------------------------------------");
        HashMap <String,Object> sheet2  = new HashMap();
        try {
            reader.hardCodeValidator(sheet,sheet2HardcodeValues);
        } catch (RuntimeException ex){

        }

        try {
            sheet2.putAll(readTable1(path));
        } catch (RuntimeException ex){

        }

        try {
            sheet2.putAll(readTable2(path));
        } catch (RuntimeException ex){

        }
        try {
            sheet2.putAll(readTable3(path));
        } catch (RuntimeException ex){

        }
        try {
            sheet2.putAll(readTable4(path));
        } catch (RuntimeException ex){

        }

        System.out.println("Sheet 2 fully validated");
        return sheet2;
    }

    public HashMap readTable1 (String path) throws IOException {

        HashMap <String,Object> table1part1  = new HashMap();
        HashMap <String,Object> table1part2  = new HashMap();
        HashMap <String,Object> table1  = new HashMap();
        HashMap <String,Object> sizeMetrix  = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(2);

        /*
        Read Table 1 of the sheet 1
        */

        //String [] headerArray1 = {"BUSINESS DIVISION","CUSTOMER NAME/CATEGORY","SEASON","STYLE NO","SAMPLE SMV"};
        //String [] headerArray2 = {"SILHOUETTE","MERCHANT NAME","GARMENT TECH NAME"};
        //int rowIndexOfHeaderStart, int columnIndexOfHeaders, String[] headerArray
        int [] rawHeaders = reader.rowHeaderValidator(sheet,2,0,sheet2Table1HeaderArray1);
        int [] rawheaders2 = reader.rowHeaderValidator(sheet,2,15,sheet2Table1HeaderArray2);
//
        table1part1 = reader.readRowHeaderTable(sheet,0, rawHeaders,2,2 );
        table1part2 = reader.readRowHeaderTable(sheet,15,rawheaders2,17,17);
        //table1.putAll(table1part1);
        //table1.putAll(table1part2);

/** ----------------------------------------------------------------------------------------------------------------------- **/

        //String [] headerArray3 = {"QTY" };
        //String [] headerArray4 = {"XS","S","M","L","XL","XXL" };
        //int rowIndexOfHeaderStart, int columnIndexOfHeaders, String[] headerArray

        int [] rawheaders3 = reader.rowHeaderValidator(sheet,3,5,sheet2SizeMetrixRows);
        int [] columnHeaders4 = reader.columnHeaderValidator(sheet, 3, sheet2SizeMetrixColumns);

//
        //int [] rawHeaders3={3,4};
        //int [] rawheaders4 ={5,7,8,9,10,11,12};

        table1part2 = reader.readColumnAndRowHeaderTable(sheet, 3, 7, columnHeaders4, 3, 4, rawheaders3);

        table1.put("cutting metrix",table1part2);

        table1.put(reader.getValue(sheet.getRow(3).getCell(13)).toString(),reader.getValue(sheet.getRow(4).getCell(13)));
        table1.put(reader.getValue(sheet.getRow(3).getCell(14)).toString(),reader.getValue(sheet.getRow(4).getCell(14)));

        table1.put(reader.getValue(sheet.getRow(2).getCell(5)).toString(),reader.getValue(sheet.getRow(5).getCell(7)));
        table1.put(reader.getValue(sheet.getRow(6).getCell(5)).toString(),reader.getValue(sheet.getRow(6).getCell(7)));
        table1.put(reader.getValue(sheet.getRow(6).getCell(14)).toString(),reader.getValue(sheet.getRow(6).getCell(17)));

        return table1;


    }

    public HashMap readTable2 (String path) throws IOException {

        HashMap <String,Object> table2part1  = new HashMap();
        HashMap <String,Object> table2part2  = new HashMap();
        HashMap <String,Object> table2  = new HashMap();
        HashMap <String,Object>conformation = new HashMap();
        HashMap <String,Object>verification = new HashMap();

        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(2);

        /*
        Read Table 1 of the sheet 1
        */

        //String [] headerArray1 = {"MARKER #","IM #","FABRIC COLOR","Fabric Face","SPECIAL NOTES"};
        //String [] headerArray2 = {"MARKER A","MARKER B","MARKER C","MARKER D","MARKER E","MARKER F"};
        //int rowIndexOfHeaderStart, int columnIndexOfHeaders, String[] headerArray

        int [] rawHeaders = reader.columnHeaderValidator(sheet, 9, sheet2FabricDetailsColumns);
        int [] rawheaders2 = reader.rowHeaderValidator(sheet,9,0,sheet2FabricDetailsRows);
//
        table2part1 = reader.readColumnAndRowHeaderTable(sheet, 9, 0, rawHeaders, 9, 13, rawheaders2);

        table2.put("Marker",table2part1);
        //----------------------------
        conformation.put(reader.getValue(sheet.getRow(14).getCell(0)).toString(),reader.getValue(sheet.getRow(14).getCell(3)));

        String temp = reader.getValue(sheet.getRow(14).getCell(10)).toString();
        String key = temp.substring(0, 4);
        String value = temp.substring( 4);
        if (temp.length() > 4 ){
            conformation.put(key, value);
        }else
            conformation.put(temp,"");
        conformation.put(reader.getValue(sheet.getRow(14).getCell(15)).toString(),reader.getValue(sheet.getRow(14).getCell(17)));

        table2.put("conformation", conformation);

        //--------------------------------
        verification .put(reader.getValue(sheet.getRow(15).getCell(0)).toString(),reader.getValue(sheet.getRow(15).getCell(3)));

        String temp1 = reader.getValue(sheet.getRow(15).getCell(10)).toString();
        String key1 = temp1.substring(0, 4);
        String value1 = temp1.substring( 4);
        if (temp1.length() > 4 ){
            verification.put(key1, value1);
        }else
            verification.put(temp1,"");

        verification.put(reader.getValue(sheet.getRow(15).getCell(15)).toString(),reader.getValue(sheet.getRow(15).getCell(17)));

        table2.put("verification", verification);



        return table2;


    }

    public HashMap readTable3 (String path) throws IOException {

        HashMap <String,Object> table3part1  = new HashMap();
        HashMap <String,Object> table3part2  = new HashMap();
        HashMap <String,Object> table3  = new HashMap();
        HashMap <String,Object>conformation = new HashMap();
        HashMap <String,Object>verification = new HashMap();

        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(2);

        /*
        Read Table 1 of the sheet 1
        */

        //String [] headerArray1 = {"SIZE","RATIO","NO OF PLIES","TOTAL CUT QTY","REQUIRED QTY"};
        //String [] headerArray2 = {"XS","S","M","L","XL" };
        //int rowIndexOfHeaderStart, int columnIndexOfHeaders, String[] headerArray

        int [] rawHeaders = reader.columnHeaderValidator(sheet, 19, sheet2CuttingInfoColumns);
        int [] rawheaders2 = reader.rowHeaderValidator(sheet,19,0,sheet2CuttingInfoRows);
//
        table3part1 = reader.readColumnAndRowHeaderTable(sheet, 19, 0, rawHeaders, 19, 23, rawheaders2);

        table3.put("cutting metrix",table3part1);
        //----------------------------

        String temp = reader.getValue(sheet.getRow(17).getCell(0)).toString();
        String key = temp.substring(0, 14);
        String value = temp.substring(14);
        if (temp.length() > 13 ){
            table3.put(key, value);
        }else
            table3.put(temp,"");


        String temp1 = reader.getValue(sheet.getRow(18).getCell(0)).toString();
        String key1 = temp1.substring(0,14);
        String value1 = temp1.substring( 14);
        if (temp1.length() > 15 ){
            table3.put(key1, value1);
        }else
            table3.put(temp1,"");

        String temp2 = reader.getValue(sheet.getRow(17).getCell(15)).toString();
        String key2 = temp2.substring(0, 10);
        String value2 = temp2.substring( 10);
        if (temp2.length() > 10 ){
            table3.put(key2, value2);
        }else
            table3.put(temp2,"");


        String temp3 = reader.getValue(sheet.getRow(18).getCell(15)).toString();
        String key3 = temp3.substring(0,10);
        String value3 = temp3.substring(10);
        if (temp3.length() >10 ){
            table3.put(key3, value3);
        }else
            table3.put(temp3,"");


        return table3;


    }

    public HashMap readTable4 (String path) throws IOException {

        HashMap <String,Object> table4 = new HashMap();
        HashMap <String,Object> cutverification = new HashMap();
        HashMap <String,Object>cpiverification = new HashMap();

        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(2);



        String temp = reader.getValue(sheet.getRow(26).getCell(0)).toString();
        String key = temp.substring(0, 13);
        String value = temp.substring(13);
        if (temp.length() > 13 ){
            cpiverification.put(key, value);
        }else
            cpiverification.put(temp,"");


        String temp1 = reader.getValue(sheet.getRow(27).getCell(0)).toString();
        String key1 = temp1.substring(0,10);
        String value1 = temp1.substring( 10);
        if (temp1.length() > 10 ){
            cpiverification.put(key1, value1);
        }else
            cpiverification.put(temp1,"");



        table4.put("CPI VERIFICATION", cpiverification);
        //------------------------------------

        String temp2 = reader.getValue(sheet.getRow(26).getCell(16)).toString();
        String key2 = temp2.substring(0, 22);
        String value2 = temp2.substring( 22);
        if (temp2.length() > 22 ){
            cutverification.put(key2, value2);
        }else
            cutverification.put(temp2,"");


        String temp3 = reader.getValue(sheet.getRow(27).getCell(16)).toString();
        String key3 = temp3.substring(0,10);
        String value3 = temp3.substring(10);
        if (temp3.length() >10 ){
            cutverification.put(key3, value3);
        }else
            cutverification.put(temp3,"");

        table4.put("Cut BANK VERIFICATION", cutverification);
        return table4;


    }

}



