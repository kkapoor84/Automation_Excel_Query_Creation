package ExcelProcessor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by Sandeep on 21 Feb 2015.
 */
public class ExcelUtil {


    public static String GetUniqueFileName()
    {
        DateFormat dateFormat = new SimpleDateFormat("dd_MMM_yyyy-HH_mm_ss");
        Date date = new Date();
        return "Insert_Queries_" + dateFormat.format(date) + ".txt";
    }

    public boolean ReadExcel2010AndCreateQuery(String excelFilePathWithFileName, int excelColumnLength, String textFilePath) {
        try {
            String textFilePathWithFileName = textFilePath + GetUniqueFileName();

            //Load excel file content into memory
            //excel processing start
            FileInputStream file = new FileInputStream(new File(excelFilePathWithFileName));

            //constructor call
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            String sheetName = sheet.getSheetName();

            //Iterate through each rows one by one
            ArrayList<String> allExcelColumnValuesArray= new ArrayList<String>();
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                for (int j = 0; j < excelColumnLength; j++) {
                    Cell cell = row.getCell(j);

                    DataFormatter df = new DataFormatter();
                    //convert row to string
                    String sc_row = df.formatCellValue(cell);
                    //System.out.println(sc_row);
                    allExcelColumnValuesArray.add(sc_row);
                }
            }
            //excel processing end...data has been read and added in array of string...suppose there are 3 rows,each of 50 columns
            //then within array ,it will hv number of rows*no of column=3*50=150 items in rowarray
            String finalQuery = createInsertQueries(allExcelColumnValuesArray, excelColumnLength, sheetName);
            //call dbInsert function to insert in database
            //DbInsert(finalQuery);
            //Write to text fille
            WriteToFile(textFilePathWithFileName, finalQuery);
            file.close();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    public String createInsertQueries(ArrayList<String> allExcelColumnValues, int excelColumnLength, String excelSheetName) throws Exception {
        try {

            String columnNamesPlaceHolder= "£££";
            String columnValuesPlaceHolder= "$$$";
            //Join column values from 0-56
            //Then append all values one by one in same from 0 to 56
                                                     // insert into tablename(colmunn1,column2) values (1,2)
            String insertQueryTemplate = String.format("INSERT INTO " + excelSheetName +  "("+ columnNamesPlaceHolder+") VALUES ( " + columnValuesPlaceHolder + " );");

            //column name reterival start
            StringBuilder columnNamesBuilder = new StringBuilder();
            //Fetch column names from ArrayList
            for (int i = 0; i < excelColumnLength; i++) {
                String currentCellValue = allExcelColumnValues.get(i);

                if (i == excelColumnLength-1) {
                    columnNamesBuilder.append(currentCellValue);
                } else {
                    columnNamesBuilder.append(currentCellValue).append(", ");
                }
            }

            /// Insert into tableName (columnName1, columnName2) values ( '$$$' );
            insertQueryTemplate = insertQueryTemplate.replace(columnNamesPlaceHolder, columnNamesBuilder.toString());
            //column name retrival finish

            //column values retrieval start
            int totalRowsInExcel = allExcelColumnValues.size() / excelColumnLength;
            int remainingRowsToProcess = totalRowsInExcel -1 ;
            StringBuilder allInsertQueriesList = new StringBuilder();
            int currentRowIndex = 1; // actually means row 2 - where values are mentioned
            for (int j = 1; j <= remainingRowsToProcess; j++) {

                int currentColumnStartIndex = excelColumnLength * currentRowIndex;
                int currentColumnEndIndex = excelColumnLength * (currentRowIndex + 1);
                String tempInsertQueryTemplate = insertQueryTemplate; /// Insert into tableName (columnName1, columnName2) values ( '$$$' );

                StringBuilder currentRowValuesBuilder = new StringBuilder();
                for (int k = currentColumnStartIndex; k < currentColumnEndIndex; k++) {
                    if (k == currentColumnEndIndex - 1) {
                        currentRowValuesBuilder.append("'").append(allExcelColumnValues.get(k)).append("' ");
                    } else {
                        currentRowValuesBuilder.append("'").append(allExcelColumnValues.get(k)).append("', ");
                    }
                }
                //before replace : Insert into tableName (columnName1, columnName2) values ( '$$$' );
                String currentRowInsertQuery = tempInsertQueryTemplate.replace(columnValuesPlaceHolder, currentRowValuesBuilder.toString());
                //after replace : Insert into tableName (columnName1, columnName2) values ( value1, value2 );

                allInsertQueriesList.append(currentRowInsertQuery);
                allInsertQueriesList.append(System.getProperty("line.separator"));
                currentRowIndex++;
                System.out.println(currentRowInsertQuery);
            }
            //column values retrieval finish

            return allInsertQueriesList.toString();


        } catch (Exception e) {
            e.printStackTrace();
            return "insert query create failed, please check the length of columns, excel file path & file to wrote path";
        }
    }

    public void WriteToFile(String path, String contentToWrite) {
        try {
            PrintStream out = new PrintStream(new FileOutputStream(path));

            out.println(contentToWrite);

            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }


    public void DbInsert(String finalInsertQueries) throws Exception {
/*
        Class.forName(DRIVER_CLASS);
        String ActualFileName = "";
        try {
            Connection conn = DriverManager.getConnection(DB_CONNECTION, DB_USER, DB_PASSWORD);
            Statement stmt = conn.createStatement();
            stmt.executeQuery(finalInsertQueries);

        } catch (Exception e) {
            System.out.println(e);
        }
*/
    }

}
