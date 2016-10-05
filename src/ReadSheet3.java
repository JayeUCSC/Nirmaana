import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

/**
 * Created by Jaye on 10/3/2016.
 */
public class ReadSheet3 {
    ExcelReader reader =  new ExcelReader();

    public HashMap finalOutput (String path) throws IOException {

        Object [][] hardcodeValues = {{35,0,"FTT 1 PERCENTAGE"},{35,2,"QUALITY PERSON NAME & SIGNATURE"},{30,7,"CHECK FINISH"},{7,15,"FTT - 2 DETAILS"}, {9,15,"FTT -2 STATUS\n" +
                "(PLEASE TICK) (ü)"}};
        HashMap <String,Object> sheet3  = new HashMap();
        sheet3.putAll(sewingInfo(path));
        sheet3.putAll(qualityInfo(path));
        sheet3.putAll(readFitDetails(path));
        sheet3.putAll(ironDetails(path));
        sheet3.putAll(ptpAndOtdDetails(path));
        return sheet3;
    }

    public HashMap sewingInfo (String path) throws IOException {

        HashMap <String,Object> table  = new HashMap();
        HashMap <String,Object> SewingTable  = new HashMap();
        HashMap <String,Object> table1part1  = new HashMap();
        HashMap <String,Object> table1part2  = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(3);
        String header = reader.getValue(sheet.getRow(7).getCell(0)).toString();

        String [] headerArray1 = {"STYLE FED DATE","STYLE FED TIME","IRON HAND OVER DATE/TIME"};
        String [] headerArray2 = {"TEAM NO","TEAM LEADER","NO OF SMO'S ALLOCATED"};

        int [] rawHeaders = reader.rowHeaderValidator(sheet,8,0,headerArray1);
        int [] rawheaders2 = reader.rowHeaderValidator(sheet,8,7,headerArray2);

        table1part1 = reader.readRowHeaderTable(sheet, 0, rawHeaders,2,2 );
        table1part2 = reader.readRowHeaderTable(sheet,7,rawheaders2,11,11);
        table.putAll(table1part1);
        table.putAll(table1part2);

        /** -------------------------------------------------------------------------------------------------------------------- **/

        String [][] headerArray3 = {{"TEAM MEMBER EPF NO"},{"TEAM MEMBER NAME"},{"SEWING START DATE"},{"SEWING START TIME"},{"SEWING FINISH DATE"},{"SEWING FINISH TIME"},
                {"TOTAL TIME DURATION"},{"OT STATUS"},{"OT HRS"}};

        int [][] columnHeaders = reader.validator(sheet,14,headerArray3 );
        HashMap table2 = reader.readComplexColumnHeaderTable(sheet, 14, columnHeaders, 15, 27);
        table.putAll(table2);

        SewingTable.put(header,table);
        return SewingTable;
    }

    public HashMap qualityInfo (String path) throws IOException {

        HashMap<String, Object> table = new HashMap();
        HashMap<String, Object> part1 = new HashMap();
        HashMap<String, Object> part2 = new HashMap();
        HashMap<String, Object> damageInfo = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(3);
        String header = reader.getValue(sheet.getRow(29).getCell(0)).toString();

        String [] Columns = {"REQUIRED","RECEIEVED","CHECK","PASS"};
        String [] Headers = {"QTY"};

        int[] columnNumbers = reader.columnHeaderValidator(sheet,30,Columns);
        int[] rowIndexes = reader.rowHeaderValidator(sheet,30,0,Headers);

        part1 =reader.readColumnAndRowHeaderTable(sheet,30, 0, columnNumbers, 31, 31, rowIndexes);
        part1.put(reader.getValue(sheet.getRow(35).getCell(0)).toString(),reader.getValue(sheet.getRow(35).getCell(1)).toString());
        part1.put(reader.getValue(sheet.getRow(35).getCell(2)).toString(),reader.getValue(sheet.getRow(35).getCell(4)).toString());

        part2.put("Date", reader.getValue(sheet.getRow(30).getCell(9)).toString());
        part2.put("Time", reader.getValue(sheet.getRow(30).getCell(15)).toString());

        part1.put(reader.getValue(sheet.getRow(30).getCell(7)).toString(), part2);

        String temp = reader.getValue(sheet.getRow(33).getCell(7)).toString();
        String key = temp.substring(0,19);
        String value = temp.substring(19);
        if (temp.length() >19 ){
            part1.put(key, value);
        }else {
            part1.put(temp, "");
        }

        table.put(header,part1);

        String header2 = reader.getValue(sheet.getRow(39).getCell(0)).toString();
        String [][] damageColumns = {{"EPF NO"},{"NAME"},{"CRITICAL"} ,{"MAJOR"},{"MINOR"},{"DEFERECT CODE"}};
        int[][] damageColumnNumbers = reader.validator(sheet, 40,damageColumns);
        damageInfo = reader.readComplexColumnHeaderTable(sheet,40,damageColumnNumbers,41,45);
        table.put(header2,damageInfo);

        return table;
    }

    public HashMap readFitDetails (String path) throws IOException {

        HashMap<String, Object> table = new HashMap();
        HashMap<String, Object> part1 = new HashMap();
        HashMap<String, Object> part2 = new HashMap();
        HashMap<String, Object> combine = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(3);
        String header = reader.getValue(sheet.getRow(7).getCell(15)).toString();

        part1.put( reader.getValue(sheet.getRow(9).getCell(17)).toString(), reader.getValue(sheet.getRow(10).getCell(17)).toString());
        part1.put( reader.getValue(sheet.getRow(9).getCell(18)).toString(), reader.getValue(sheet.getRow(10).getCell(18)).toString());
        combine.put(reader.getValue(sheet.getRow(9).getCell(15)).toString(),part1);

        String [] rowHeaders = {"REASONS FOR FAILURES","REJECTED QTY"};
        int [] rows = reader.columnHeaderValidator(sheet,11,rowHeaders);
        part2 = reader.readColumnHeaderTable(sheet,11,rows,12,18);
        combine.putAll(part2);

        table.put(header,combine);
        return table;
    }

    public HashMap ironDetails (String path) throws IOException {

        HashMap<String, Object> table = new HashMap();
        HashMap<String, Object> combine = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(3);

        String[] headers = {"IRON EPF", "IRONER NAME", "IRON START DATE", "IRON START TIME", "IRON FINISH DATE", "IRON FINISH TIME", "IRON QTY"};
        int[] rowHeaders = reader.rowHeaderValidator(sheet,19,15,headers);
        combine = reader.readRowHeaderTable(sheet, 15, rowHeaders, 17, 17);

        String temp = reader.getValue(sheet.getRow(26).getCell(15)).toString() ;
        String key = temp.substring(0,15);
        String value = temp.substring(15);
        if (temp.length() >15 ){
            combine.put(key, value);
        }else {
            combine.put(temp, "");
        }

        table.put("Iron Details", combine);
        return table;
    }

    public HashMap ptpAndOtdDetails (String path) throws IOException {

        HashMap<String, Object> ptp = new HashMap();
        HashMap<String, Object> odt = new HashMap();
        HashMap<String, Object> ptpStatus = new HashMap();
        HashMap<String, Object> otdStatus = new HashMap();
        HashMap<String, Object> failureReasons = new HashMap();
        HashMap<String, Object> combine = new HashMap();
        HashMap<String, Object> table = new HashMap();
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(3);

        ptpStatus.put(reader.getValue(sheet.getRow(3).getCell(22)).toString(), reader.getValue(sheet.getRow(4).getCell(22)).toString());
        ptpStatus.put(reader.getValue(sheet.getRow(3).getCell(23)).toString(), reader.getValue(sheet.getRow(4).getCell(23)).toString());
        ptp.put(reader.getValue(sheet.getRow(3).getCell(20)).toString(), ptpStatus);

        otdStatus.put(reader.getValue(sheet.getRow(5).getCell(22)).toString(), reader.getValue(sheet.getRow(6).getCell(22)).toString());
        otdStatus.put(reader.getValue(sheet.getRow(5).getCell(23)).toString(), reader.getValue(sheet.getRow(6).getCell(23)).toString());
        odt.put(reader.getValue(sheet.getRow(5).getCell(20)).toString(), otdStatus);

        combine.put(reader.getValue(sheet.getRow(4).getCell(22)).toString(), odt);
        combine.put(reader.getValue(sheet.getRow(2).getCell(20)).toString(), ptp);

        String[] rowHeaders = {"FABRIC DAMAGE (HOLE/RUN)", "FABRIC SHADING", "MEASUREMENT - TO TO SPEC",
                "CONSTUCTION - BORKEN/DROP/SKIP STITCH", "CONSTRUCTION - RAW EDGES / FRAYED SEAMS",
                "CONSTRUCTION - OPEN/ DELAMINATED SEAM", "CONSTRUCTION - OVERRUN STITCHES",
                "CONSTRUCTION - UNEVEN / WAVY STITCHING", "CONSTRUCTION - PUCKERING / PLEATING",
                "CONSTRUCTION - TWISTED ROPING / UNEVEN HEM", "CONSTRUCTION NOT AS SPECIFIED",
                "CONSTRUCTION - MISCELLANEOUS  DEFECTS", "TRIM - BROKEN/ TRIM INOPERABLE",
                "EMBELLISHMENT - UNRAVELLING", "EMBELLISHMENT - PEELING / DELAMINATION PRINT",
                "EMBELLISHMENT - POOR COVERAGE/ POOR REGISTRATION/ CRACKING", "EMBELLISHMENT - OUT OF TOLERANCE",
                "LABEL - MISSING / INCORRECT / IN COMPLETE",
                "APPEARANCE - SOIL / DIRT MARK", "APPEARANCE - UNATTACHED / UNTRIMMED THREADS", "OTHER ISSUE"};
        int[] headerIndexes = reader.rowHeaderValidator(sheet, 8, 20, rowHeaders);
        failureReasons = reader.readRowHeaderTable(sheet, 20, headerIndexes, 23, 23);

        combine.put(reader.getValue(sheet.getRow(7).getCell(20)).toString(), failureReasons);
        table.put(reader.getValue(sheet.getRow(0).getCell(20)).toString(), combine);
        return table;
    }
}
