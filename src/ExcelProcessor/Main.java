package ExcelProcessor;

/**
 * Created by Sandeep on 22 Feb 2015.
 */
public class Main {

	//Startup function for the program//
    public static void main(String[] args) {

        String excelFilePathWithFileName = "F:\\JavaProjects\\ExcelPoc\\ExcelFiles\\excel_data_2010.xlsx";
        String textFilePath = "F:\\JavaProjects\\ExcelPoc\\ExcelFiles\\"; //Only file path
        int excelColumnLength = 57;

        //Create object
        ExcelUtil excelUtilInstance = new ExcelUtil();

        //Call method ReadExcel2010AndCreateQuery for excel 2010 ( .xlsx format )
        excelUtilInstance.ReadExcel2010AndCreateQuery(excelFilePathWithFileName, excelColumnLength, textFilePath);
    }
}
