/**
 * Created by Jaye on 10/6/2016.
 */
public interface Configuration {

    /** ----------------------------------------------- sheet1 ---------------------------------------------------------------- **/

    public Object [][] sheet1HardcodeValues = {{6,9,"Total"},{6,10,"Test Cut Requirement"}};
    String [] sheet1Table1HeaderArray1 = {"Business Division","Style No","Season","Style Name","Silhouette"};
    String [] sheet1Table1HeaderArray2 = {"Customer Name/Category","Sample Type","Merchant Name","Garment Tech Name"};

    String [] sheet1sizeMetrixColumns = {"XS","S","M","L","XL","XXL"};
    String [] sheet1sizeMetrixHeaders = {"QTY","Aditional QTY"};

    String [][] sheet1table2Headers = {{"ITEM"},{"IM"},{"Sub IM"},{"ITEM COLOR"},{"Sub color"},{"CW"},{"WIDTH/SIZE"},{"ITEM QTY"},{"Fabric Face Side"},
            {"USAGE"},{"SUPPLIER"},{"DESCREPTION"},{"COMMENTS AND REMARKS"}};

    /** ----------------------------------------------- sheet2 ---------------------------------------------------------------- **/

    Object [][] sheet2HardcodeValues = {{2,5,"SAMPLE TYPE"},{6,5,"EARLY COMPETION TIME"},{3,13,"Add. Cuts"}, {3,14,"Total"},{6,14,"SAMPLE PLAN END DATE & TIME"},
            {14,0,"Confirmation From Stores"},{14,10,"DATE",14,15,"TIME:"},{15,0,"Verification (G.Tech)"},{15,10,"DATE"},{15,15,"TIME:"},
            {17,0,"CUTTING TEAM :"},{17,15,"CUT DATE :"},{18,0,"CUTTER/S NAME :"},{18,15,"CUT TIME :"} ,
            {26,0,"CPI DONE BY :"},{26,16,"VERIFICATION DONE BY :"},{27,0,"COMMENTS :"},{27,16,"COMMENTS :"}};

    String [] sheet2Table1HeaderArray1 = {"BUSINESS DIVISION","CUSTOMER NAME/CATEGORY","SEASON","STYLE NO","SAMPLE SMV"};
    String [] sheet2Table1HeaderArray2 = {"SILHOUETTE","MERCHANT NAME","GARMENT TECH NAME"};

    String [] sheet2SizeMetrixColumns = {"XS","S","M","L","XL","XXL" };
    String [] sheet2SizeMetrixRows = {"QTY" };

    String [] sheet2FabricDetailsColumns = {"MARKER A","MARKER B","MARKER C","MARKER D","MARKER E","MARKER F"};
    String [] sheet2FabricDetailsRows = {"MARKER #","IM #","FABRIC COLOR","Fabric Face","SPECIAL NOTES"};

    String [] sheet2CuttingInfoColumns = {"XS","S","M","L","XL" };
    String [] sheet2CuttingInfoRows = {"SIZE","RATIO","NO OF PLIES","TOTAL CUT QTY","REQUIRED QTY"};

    /** ----------------------------------------------- sheet3 ---------------------------------------------------------------- **/

    Object [][] sheet3HardcodeValues = {{7,0,"SEWING INFORMATION"},{29,0,"QUALITY INFORMATION"},{35,0,"FTT 1 PERCENTAGE"},{35,2,"QUALITY PERSON NAME & SIGNATURE"},{30,7,"CHECK FINISH"},{39,0,"DAMAGE STATUS"},{7,15,"FTT - 2 DETAILS"}, {9,15,"FTT -2 STATUS\n" +
            "(PLEASE TICK) (ü)"},{0,20,"PTP & OTD DETAILS"},{3,20,"PTP"},{5,20,"OTD \n" +
            "(PLEASE TICK)"},{7,20,"FAILURE REASONS"}};

    String [] sheet3SewingInfoHeaderArray1 = {"STYLE FED DATE","STYLE FED TIME","IRON HAND OVER DATE/TIME"};
    String [] sheet3SewingInfoHeaderArray2 = {"TEAM NO","TEAM LEADER","NO OF SMO'S ALLOCATED"};
    String [][] sheet3SewingInfoHeaderArray3 = {{"TEAM MEMBER EPF NO"},{"TEAM MEMBER NAME"},{"SEWING START DATE"},{"SEWING START TIME"},{"SEWING FINISH DATE"},{"SEWING FINISH TIME"},
            {"TOTAL TIME DURATION"},{"OT STATUS"},{"OT HRS"}};

    String [] sheet3QualityInfoColumns = {"REQUIRED","RECEIEVED","CHECK","PASS"};
    String [] sheet3QualityInforow = {"QTY"};
    String [][] sheet3DamageColumns = {{"EPF NO"},{"NAME"},{"CRITICAL"} ,{"MAJOR"},{"MINOR"},{"DEFERECT CODE"}};

    String [] sheet3FttDetailsRowHeaders = {"REASONS FOR FAILURES","REJECTED QTY"};
    String[] sheet3IronHeaders = {"IRON EPF", "IRONER NAME", "IRON START DATE", "IRON START TIME", "IRON FINISH DATE", "IRON FINISH TIME", "IRON QTY"};

    String[] sheet3PtpAndOtdrowHeaders = {"FABRIC DAMAGE (HOLE/RUN)", "FABRIC SHADING", "MEASUREMENT - TO TO SPEC",
            "CONSTUCTION - BORKEN/DROP/SKIP STITCH", "CONSTRUCTION - RAW EDGES / FRAYED SEAMS",
            "CONSTRUCTION - OPEN/ DELAMINATED SEAM", "CONSTRUCTION - OVERRUN STITCHES",
            "CONSTRUCTION - UNEVEN / WAVY STITCHING", "CONSTRUCTION - PUCKERING / PLEATING",
            "CONSTRUCTION - TWISTED ROPING / UNEVEN HEM", "CONSTRUCTION NOT AS SPECIFIED",
            "CONSTRUCTION - MISCELLANEOUS  DEFECTS", "TRIM - BROKEN/ TRIM INOPERABLE",
            "EMBELLISHMENT - UNRAVELLING", "EMBELLISHMENT - PEELING / DELAMINATION PRINT",
            "EMBELLISHMENT - POOR COVERAGE/ POOR REGISTRATION/ CRACKING", "EMBELLISHMENT - OUT OF TOLERANCE",
            "LABEL - MISSING / INCORRECT / IN COMPLETE",
            "APPEARANCE - SOIL / DIRT MARK", "APPEARANCE - UNATTACHED / UNTRIMMED THREADS", "OTHER ISSUE"};

    /** ----------------------------------------------- sheet4 ---------------------------------------------------------------- **/

    Object [][] sheet4HardcodeValues = {{10,0,"TRIM STORES"}};

    String [] sheet4Table1ColumnHeaders = {"MARKER A","MARKET B","MARKER C","MARKER D","MARKER E","MARKER F"};
    String [] sheet4Table1RowHeaders = {"IM #","FABRIC COLOR"};

    String [] sheet4TrimStoresColumnHeaders = {"1","2","3","4","5","6","7"};
    String [] sheet4TrimStoresRowHeaders = {"IM#","COLOUR","SIZE","DESCRIPTION","QTY"};

    String[] sheet4LastTableColumnHeaders = {"1", "2", "3", "4", "5", "6", "7"};
    String[] sheet4LastTableRowHeaders = {"IM#","COLOUR","SIZE/WIDTH","DESCRIPTION","QTY"};
}
