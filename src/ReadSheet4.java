import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * Created by Jaye on 10/4/2016.
 */
public class ReadSheet4 {
    ExcelReader reader =  new ExcelReader();

    public HashMap finalOutput (String path) throws IOException {

        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(4);
        Object [][] hardcodeValues = {{10,0,"TRIM STORES"}};
        reader.hardCodeValidator(sheet,hardcodeValues);
        System.out.println("Sheet 4 fully validated");

        HashMap <String,Object> sheet3  = new HashMap();
        sheet3.putAll(readTable1(path));
        sheet3.putAll(trimStores(path));
        sheet3.putAll(lastTable(path));
        return sheet3;
    }

    public HashMap readTable1 (String path) throws IOException {

        HashMap<String, Object> table = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(4);

        String [] columnHeaders = {"MARKER A","MARKET B","MARKER C","MARKER D","MARKER E","MARKER F"};
        String [] rowHeaders = {"IM #","FABRIC COLOR"};

        int [] columnIndexes = reader.columnHeaderValidator(sheet,7,columnHeaders);
        int [] rowIndexes = reader.rowHeaderValidator(sheet,8,0,rowHeaders);

        table = reader.readColumnAndRowHeaderTable(sheet,7,0,columnIndexes,8,9,rowIndexes);
        return table;

    }

    public HashMap trimStores (String path) throws IOException {

        HashMap<String, Object> table1 = new HashMap();
        HashMap<String, Object> table = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(4);

        String [] columnHeaders = {"1","2","3","4","5","6","7"};
        String [] rowHeaders = {"IM#","COLOUR","SIZE","DESCRIPTION","QTY"};

        int [] columnIndexes = reader.columnHeaderValidator(sheet,12,columnHeaders);
        int [] rowIndexes = reader.rowHeaderValidator(sheet,12,0,rowHeaders);

        table1 = reader.readColumnAndRowHeaderTable(sheet,12,0,columnIndexes,13,17,rowIndexes);
        table.put(reader.getValue(sheet.getRow(10).getCell(0)).toString(),table1);
        return table;

    }

    public HashMap lastTable (String path) throws IOException {

        HashMap<String, Object> table1 = new HashMap();
        HashMap<String, Object> table = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(4);

        String[] columnHeaders = {"1", "2", "3", "4", "5", "6", "7"};
        String[] rowHeaders = {"IM#","COLOUR","SIZE/WIDTH","DESCRIPTION","QTY"};

        int[] columnIndexes = reader.columnHeaderValidator(sheet,24,columnHeaders);
        int[] rowIndexes = reader.rowHeaderValidator(sheet,25,0,rowHeaders);

        table1 = reader.readColumnAndRowHeaderTable(sheet,24,0,columnIndexes,25,29,rowIndexes);
        table.put("Pre sewing trims",table1);

        return table;
    }
}
