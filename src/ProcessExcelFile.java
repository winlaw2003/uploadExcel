import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;

public class ProcessExcelFile {

    public static final String EXCEL_FILE = "./Input_Form_1040nr.xlsm";

    public static void main(String[] args) throws InvalidFormatException, IOException {

        Workbook workbook;
        Sheet sheet;

        ExcelWorkbook excelWorkbook = new ExcelWorkbook();

        workbook = excelWorkbook.getWorkbook(EXCEL_FILE);
        sheet = excelWorkbook.getSheet(0);

        System.out.println("Workbook has " + excelWorkbook.getTotalWorksheets() + " Sheets : ");

//        excelWorkbook.ListSheetNames();


        // create a ArrayList String type
        // and Initialize an ArrayList with add()
        ArrayList<String> sheetnames = new ArrayList<String>() {
            {
                for (Sheet sheet : workbook) {
                    System.out.println("=> " + sheet.getSheetName());
                    add(sheet.getSheetName());
                }
            }
        };
        System.out.println("ArrayList : " + sheetnames);



        excelWorkbook.CloseWorkbook();


    }
}
