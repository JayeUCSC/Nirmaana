
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;

/**
 * Created by Amila on 2016-09-27.
 */
public class ExcelReader {

    public void a(){
        System.out.println("ss");
    }

    public HashMap readColumnHeaderTable(XSSFSheet workSheet, int rowIndexOfHeaderStart, int[] headerIndexes, int rowIndexOfStartCell, int rowIndexOfEndCell) {

        HashMap outerHashmap = new HashMap();
        Arrays.sort(headerIndexes);
        int columnIndexOfHeaderStart = headerIndexes[0];
        int columnIndexOfHeaderEnd = headerIndexes[(headerIndexes.length) - 1];
        String[] header = new String[(columnIndexOfHeaderEnd - columnIndexOfHeaderStart) + 1];
        XSSFRow headerRow = workSheet.getRow(rowIndexOfHeaderStart);

        if (headerRow.getLastCellNum() <= columnIndexOfHeaderEnd) {
            throw new RuntimeException(columnIndexOfHeaderEnd + " is not in range");
        }

        for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

            if (Arrays.binarySearch(headerIndexes, j) >= 0) {
                XSSFCell cell = headerRow.getCell(j);
                header[j - columnIndexOfHeaderStart] = getValue(cell).toString();
            }
        }

        for (int i = rowIndexOfStartCell; i <= rowIndexOfEndCell; i++) {
            XSSFRow row = workSheet.getRow(i);
            HashMap innerHashmap = new HashMap();

            for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

                if (Arrays.binarySearch(headerIndexes, j) >= 0) {
                    XSSFCell cell = row.getCell(j);
                    String value = getValue(cell).toString();
                    innerHashmap.put(header[j - columnIndexOfHeaderStart], value);
                }
            }
            outerHashmap.put(i + 1, innerHashmap);
        }
        return outerHashmap;
    }

/*
    public HashMap readColumnHeaderTable(XSSFSheet workSheet, int rowIndexOfHeaderStart, int[] headerIndexes, int rowIndexOfStartCell, int rowIndexOfEndCell,int rowIndexOfMainHeader) {

        HashMap outerHashmap = new HashMap();
        Arrays.sort(headerIndexes);
        int columnIndexOfHeaderStart = headerIndexes[0];
        int columnIndexOfHeaderEnd = headerIndexes[(headerIndexes.length) - 1];
        String[] header = new String[(columnIndexOfHeaderEnd - columnIndexOfHeaderStart) + 1];
        XSSFRow headerRow = workSheet.getRow(rowIndexOfHeaderStart);

        String mainHeader = workSheet.getRow(rowIndexOfMainHeader).getCell(headerIndexes[0]).toString();
        if (headerRow.getLastCellNum() <= columnIndexOfHeaderEnd) {
            throw new RuntimeException(columnIndexOfHeaderEnd + " is not in range");
        }

        for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

            if (Arrays.binarySearch(headerIndexes, j) >= 0) {
                XSSFCell cell = headerRow.getCell(j);
                header[j - columnIndexOfHeaderStart] = getValue(cell).toString();
            }
        }

        for (int i = rowIndexOfStartCell; i <= rowIndexOfEndCell; i++) {
            XSSFRow row = workSheet.getRow(i);
            HashMap innerHashmap = new HashMap();

            for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

                if (Arrays.binarySearch(headerIndexes, j) >= 0) {
                    XSSFCell cell = row.getCell(j);
                    String value = getValue(cell).toString();
                    innerHashmap.put(header[j - columnIndexOfHeaderStart], value);
                }
            }
            outerHashmap.put(mainHeader, innerHashmap);
        }
        return outerHashmap;
    }

*/

    public HashMap readColumnHeaderTable(XSSFSheet workSheet,int rowIndexOfHeaderStart, int columnIndexOfHeaderStart,  int[] headerIndexes, int rowIndexOfStartCell, int rowIndexOfEndCell,int[] rowIndexes) {

        HashMap outerHashmap = new HashMap();
        Arrays.sort(headerIndexes);
        int columnIndexOfHeaderEnd = headerIndexes[(headerIndexes.length)-1];
        String[] header = new String[(columnIndexOfHeaderEnd - columnIndexOfHeaderStart)+1];
        XSSFRow headerRow = workSheet.getRow(rowIndexOfHeaderStart);
        System.out.println(headerRow.getLastCellNum()+"**********************"+columnIndexOfHeaderEnd);
        if(headerRow.getLastCellNum()<=columnIndexOfHeaderEnd){
            throw new RuntimeException(columnIndexOfHeaderEnd+" is not in range");
        }

        for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

            if (Arrays.binarySearch(headerIndexes, j) >= 0) {

                XSSFCell cell = headerRow.getCell(j);
                header[j-columnIndexOfHeaderStart] = getValue(cell).toString();
                System.out.println(0 + " " + j);
                //System.out.println(header[j]);
            }
        }
        System.out.println("#############################" + rowIndexOfStartCell);

        for (int i = rowIndexOfStartCell; i <= rowIndexOfEndCell; i++) {
            XSSFRow row = workSheet.getRow(i);
            HashMap innerHashmap = new HashMap();

            if (Arrays.binarySearch(rowIndexes, i) >= 0) {
                for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

                    if (Arrays.binarySearch(headerIndexes, j) >= 0) {
                        XSSFCell cell = row.getCell(j);
                        String value = getValue(cell).toString();
                        innerHashmap.put(header[j - columnIndexOfHeaderStart], value);
                    }
                }
                outerHashmap.put(i + 1, innerHashmap);
            }
        }
        return outerHashmap;
    }

    public HashMap readRowHeaderTable(XSSFSheet workSheet, int columnIndexOfHeaderStart, int[] headerIndexes, int columnIndexOfStartCell, int columnIndexOfEndCell) {

        HashMap outerHashmap = new HashMap();
        Arrays.sort(headerIndexes);
        int rowIndexOfHeaderStart = headerIndexes[0];
        int rowIndexOfHeaderEnd = headerIndexes[(headerIndexes.length) - 1];
        String[] header = new String[(rowIndexOfHeaderEnd - rowIndexOfHeaderStart) + 1];

        if (workSheet.getLastRowNum() < rowIndexOfHeaderEnd) {
            throw new RuntimeException(rowIndexOfHeaderEnd + " is not in range");
        }

        for (int i = rowIndexOfHeaderStart; i <= rowIndexOfHeaderEnd; i++) {

            if (Arrays.binarySearch(headerIndexes, i) >= 0) {
                XSSFRow headerRow = workSheet.getRow(i);
                XSSFCell cell = headerRow.getCell(columnIndexOfHeaderStart);
                header[i - rowIndexOfHeaderStart] = getValue(cell).toString();
            }
        }

        for (int j = columnIndexOfStartCell; j <= columnIndexOfEndCell; j++) {
            HashMap innerHashmap = new HashMap();

            for (int i = rowIndexOfHeaderStart; i <= rowIndexOfHeaderEnd; i++) {

                if (Arrays.binarySearch(headerIndexes, i) >= 0) {
                    XSSFRow headerRow = workSheet.getRow(i);
                    XSSFCell cell = headerRow.getCell(j);
                    String value = getValue(cell).toString();
                    innerHashmap.put(header[i - rowIndexOfHeaderStart], value);
                }
            }
            outerHashmap.put(j + 1, innerHashmap);
        }
        return outerHashmap;
    }

    /**
     * @param workSheet
     * @param rowIndexOfHeaderStart
     * @param headerArray
     * @param rowIndexOfStartCell
     * @param rowIndexOfEndCell
     * @return HashMap
     */
    public HashMap readComplexColumnHeaderTable(XSSFSheet workSheet, int rowIndexOfHeaderStart,String[][] headerArray, int rowIndexOfStartCell, int rowIndexOfEndCell) {

        int[][] headerIndexes = validator(workSheet,rowIndexOfHeaderStart,headerArray);
        HashMap outerHashmap = new HashMap();
        XSSFRow headerRow1 = workSheet.getRow(rowIndexOfHeaderStart);
        XSSFRow headerRow2 = workSheet.getRow(rowIndexOfHeaderStart + 1);

        for (int i = rowIndexOfStartCell; i < rowIndexOfEndCell; i++) {
            XSSFRow dataRow = workSheet.getRow(i);
            HashMap innerHashMap = new HashMap();

            for (int[] innerHeaderIndexes : headerIndexes) {

                if (headerRow1.getLastCellNum() <= innerHeaderIndexes[0]) {
                    throw new RuntimeException(innerHeaderIndexes[0] + " is not in range");
                }

                if (innerHeaderIndexes.length == 1) {
                    XSSFCell headerCell1 = headerRow1.getCell(innerHeaderIndexes[0]);
                    XSSFCell dataCell = dataRow.getCell(innerHeaderIndexes[0]);
                    String headerValue1 = getValue(headerCell1).toString();
                    String dataValue = getValue(dataCell).toString();

                    innerHashMap.put(headerValue1, dataValue);

                } else if (innerHeaderIndexes.length > 1) {
                    XSSFCell headerCell1 = headerRow1.getCell(innerHeaderIndexes[0]);
                    String headerValue1 = getValue(headerCell1).toString();
                    HashMap innerHashMap2 = new HashMap();
                    Arrays.sort(innerHeaderIndexes);
                    for (int innerHeaderIndex : innerHeaderIndexes) {

                        if (headerRow1.getLastCellNum() <= innerHeaderIndex) {
                            throw new RuntimeException(innerHeaderIndex + " is not in range");
                        }

                        XSSFCell headerCell2 = headerRow2.getCell(innerHeaderIndex);
                        XSSFCell dataCell = dataRow.getCell(innerHeaderIndex);
                        String headerValue2 = getValue(headerCell2).toString();
                        String dataValue = getValue(dataCell).toString();

                        innerHashMap2.put(headerValue2, dataValue);
                    }
                    innerHashMap.put(headerValue1, innerHashMap2);
                }
            }
            outerHashmap.put(i + 1, innerHashMap);

        }
        System.out.println("--- Return HashMap ---");
        return outerHashmap;
    }

    /**
     *
     * @param workSheet
     * @param rowIndexOfHeaderStart
     * @param headerIndexes
     * @param rowIndexOfStartCell
     * @param rowIndexOfEndCell
     * @return
     */

    public HashMap readComplexColumnHeaderTable(XSSFSheet workSheet, int rowIndexOfHeaderStart,int[][] headerIndexes, int rowIndexOfStartCell, int rowIndexOfEndCell) {

        HashMap outerHashmap = new HashMap();
        XSSFRow headerRow1 = workSheet.getRow(rowIndexOfHeaderStart);
        XSSFRow headerRow2 = workSheet.getRow(rowIndexOfHeaderStart + 1);

        for (int i = rowIndexOfStartCell; i < rowIndexOfEndCell; i++) {
            XSSFRow dataRow = workSheet.getRow(i);
            HashMap innerHashMap = new HashMap();

            for (int[] innerHeaderIndexes : headerIndexes) {

                if (headerRow1.getLastCellNum() <= innerHeaderIndexes[0]) {
                    throw new RuntimeException(innerHeaderIndexes[0] + " is not in range");
                }

                if (innerHeaderIndexes.length == 1) {
                    XSSFCell headerCell1 = headerRow1.getCell(innerHeaderIndexes[0]);
                    XSSFCell dataCell = dataRow.getCell(innerHeaderIndexes[0]);
                    String headerValue1 = getValue(headerCell1).toString();
                    String dataValue = getValue(dataCell).toString();

                    innerHashMap.put(headerValue1, dataValue);

                } else if (innerHeaderIndexes.length > 1) {
                    XSSFCell headerCell1 = headerRow1.getCell(innerHeaderIndexes[0]);
                    String headerValue1 = getValue(headerCell1).toString();
                    HashMap innerHashMap2 = new HashMap();
                    Arrays.sort(innerHeaderIndexes);
                    for (int innerHeaderIndex : innerHeaderIndexes) {

                        if (headerRow1.getLastCellNum() <= innerHeaderIndex) {
                            throw new RuntimeException(innerHeaderIndex + " is not in range");
                        }

                        XSSFCell headerCell2 = headerRow2.getCell(innerHeaderIndex);
                        XSSFCell dataCell = dataRow.getCell(innerHeaderIndex);
                        String headerValue2 = getValue(headerCell2).toString();
                        String dataValue = getValue(dataCell).toString();

                        innerHashMap2.put(headerValue2, dataValue);
                    }
                    innerHashMap.put(headerValue1, innerHashMap2);
                }
            }
            outerHashmap.put(i + 1, innerHashMap);

        }
        System.out.println("--- Return HashMap ---");

        return outerHashmap;
    }

    public HashMap readColumnAndRowHeaderTable(XSSFSheet workSheet,int rowIndexOfHeaderStart, int columnIndexOfHeaderStart,  int[] columnHeaderIndexes, int rowIndexOfStartCell, int rowIndexOfEndCell,int[] rowHeaderIndexes) {

        HashMap outerHashmap = new HashMap();
        Arrays.sort(columnHeaderIndexes);
        Arrays.sort(rowHeaderIndexes);
        int columnIndexOfHeaderEnd = columnHeaderIndexes[(columnHeaderIndexes.length)-1];
        String[] columnHeaders = new String[(columnIndexOfHeaderEnd - columnIndexOfHeaderStart)+1];
        String[] rowHeaders = new String[(columnIndexOfHeaderEnd - columnIndexOfHeaderStart)+1];

        XSSFRow headerRow = workSheet.getRow(rowIndexOfHeaderStart);
        XSSFCell headerCell;

        if(headerRow.getLastCellNum()<=columnIndexOfHeaderEnd){
            throw new RuntimeException(columnIndexOfHeaderEnd+" is not in range");
        }

        for (int j = columnIndexOfHeaderStart; j <= columnIndexOfHeaderEnd; j++) {

            if (Arrays.binarySearch(columnHeaderIndexes, j) >= 0) {

                XSSFCell cell = headerRow.getCell(j);
                columnHeaders[j-columnIndexOfHeaderStart] = getValue(cell).toString();

            }
        }

        for (int j = rowIndexOfHeaderStart; j <= rowIndexOfEndCell; j++) {

            if (Arrays.binarySearch(rowHeaderIndexes, j) >= 0) {

                headerCell  = workSheet.getRow(j).getCell(columnIndexOfHeaderStart);
                rowHeaders[j-rowIndexOfHeaderStart] = getValue(headerCell).toString();

            }
        }



        for (int i = columnIndexOfHeaderStart; i <= columnIndexOfHeaderEnd; i++) {

            HashMap innerHashmap = new HashMap();

            if (Arrays.binarySearch(columnHeaderIndexes, i) >= 0) {
                for (int j = rowIndexOfStartCell; j <= rowIndexOfEndCell; j++) {
                    XSSFRow row = workSheet.getRow(j);

                    if (Arrays.binarySearch(rowHeaderIndexes, j) >= 0) {
                        XSSFCell cell = row.getCell(i);
                        String value = getValue(cell).toString();
                        System.out.println(value);
                        innerHashmap.put(rowHeaders[j - rowIndexOfHeaderStart], value);
                    }
                }
                outerHashmap.put(columnHeaders[i - columnIndexOfHeaderStart], innerHashmap);
            }
        }
        return outerHashmap;
    }

    /**
     * @param cell
     * @return Object
     */
    public static Object getValue(XSSFCell cell) {
        Object result = "";
        if (cell != null) {
            int cellType = cell.getCellType();
            //System.out.println(cell.getRowIndex() + "  " + cell.getColumnIndex() + " -->" + cell.getCellType());//+ "  "+ cell.getStringCellValue());

            switch (cellType) {
                case Cell.CELL_TYPE_NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
                        result = df.format(date);

                    } else {
                        result = new DataFormatter().formatCellValue(cell);
                    }
                    //System.out.println(cell.getColumnIndex() + "  " + cell.getRowIndex() + "-->" + result.toString());
                    break;

                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue();
                    break;

                case Cell.CELL_TYPE_FORMULA:
                    result = cell.getRawValue();
                    break;

                case Cell.CELL_TYPE_BLANK:
                    break;

                case Cell.CELL_TYPE_BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;

                case Cell.CELL_TYPE_ERROR:
                    break;

                default:
                    throw new RuntimeException("Cell Type miss match");

            }

        }
        return result;
    }


    public static int[][] validator(XSSFSheet workSheet, int rowIndexOfHeaderStart, String[][] headerArray) {
        int[][] result = new int[headerArray.length][];
        boolean validation = true;

        XSSFRow headerRow1 = workSheet.getRow(rowIndexOfHeaderStart);
        XSSFRow headerRow2 = workSheet.getRow(rowIndexOfHeaderStart + 1);

        for (int i = 0; i < headerArray.length; i++) {
            boolean status = false;
            String[] innerHeaderArray = headerArray[i];

            if (innerHeaderArray.length == 1) {
                int[] indexs = new int[innerHeaderArray.length];
                for (int j = 0; j < headerRow1.getLastCellNum(); j++) {
                    XSSFCell cell = headerRow1.getCell(j);
                    if (innerHeaderArray[0].equals(getValue(cell).toString().trim())) {
                        indexs[0] = cell.getColumnIndex();
                        result[i] = indexs;
                        status = true;
                        break;
                    }
                }
            } else if (innerHeaderArray.length > 1) {
                int[] indexs = new int[innerHeaderArray.length - 1];
                for (int j = 0; j < headerRow1.getLastCellNum(); j++) {
                    XSSFCell cell = headerRow1.getCell(j);
                    if (innerHeaderArray[0].equals(getValue(cell).toString().trim())) {
                        int counter = 0;
                        for(int z = 0; z<innerHeaderArray.length-1; z++ ) {
                            String innerHeaderElemant = innerHeaderArray[z + 1];

                            for (int k = 0; k <headerRow2.getLastCellNum(); k++) {
                                XSSFCell innerCell = headerRow2.getCell(j + k);
                                if (innerHeaderElemant.equals(getValue(innerCell).toString())) {
                                    indexs[z] = innerCell.getColumnIndex();
                                    counter++;
                                    break;
                                }
                            }
                        }

                        if (counter == innerHeaderArray.length - 1) {
                            status = true;
                        }
                        result[i] = indexs;
                        break;
                    }
                }
            }
            if (status) {
                continue;
            } else {
                System.out.println(headerArray[i][0] + " column is mismatch with Excel columns");
                validation = false;
                //hti@lit123
                //itu@volta#2

            }
        }

        if (validation) {
            System.out.println("--- Headers are successfully validated ---");
            return result;
        } else {
            throw new RuntimeException("Excel columns are mismatch");
        }

    }

    public static int[] rowHeaderValidator(XSSFSheet workSheet, int rowIndexOfHeaderStart, int columnIndexOfHeaders, String[] headerArray) {
        int[] result = new int[headerArray.length];
        boolean validation = true;

        for (int i = 0;i<headerArray.length; i++) {
            String value = headerArray[i];
            boolean status = false;
            for (int j = rowIndexOfHeaderStart; j < workSheet.getLastRowNum(); j++) {
                XSSFRow row = workSheet.getRow(j);
                XSSFCell cell = row.getCell(columnIndexOfHeaders);
                if ((getValue(cell).toString()== null ? "" : getValue(cell).toString()).toString().trim().equals(value)){
                    result [i] = j;
                    status = true;
                    break;
                }

            }
            if (status) {
                continue;
            } else {
                System.out.println(headerArray[i] + " Row is mismatch with Excel columns");
                validation = false;
            }
        }
        if (validation) {
            System.out.println("--- Headers are successfully validated ---");
            return result;
        } else {
            throw new RuntimeException("Excel Rows are mismatch");
        }

    }

    public static int[] columnHeaderValidator(XSSFSheet workSheet, int rowIndexOfHeaderStart, String[] headerArray) {
        int[] result = new int[headerArray.length];
        boolean validation = true;
        XSSFRow row = workSheet.getRow(rowIndexOfHeaderStart);
        for (int i = 0;i<headerArray.length; i++) {
            String value = headerArray[i];
            boolean status = false;

            for (int j = 0; j < row.getLastCellNum(); j++) {
                XSSFCell cell= row.getCell(j);
                if ((getValue(cell).toString()== null ? "" : getValue(cell).toString()).toString().trim().equals(value)){
                    result [i] = j;
                    status = true;
                }

            }
            if (status) {
                continue;
            } else {
                System.out.println(headerArray[i] + " column is mismatch with Excel columns");
                validation = false;
            }
        }
        if (validation) {
            System.out.println("--- Headers are successfully validated ---");
            return result;
        } else {
            throw new RuntimeException("Excel columns are mismatch");
        }

    }

}

