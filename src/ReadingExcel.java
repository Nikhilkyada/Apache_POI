import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadingExcel {

    public static void main(String arg[]) {

        //file path stored in variable
        String excelFilePath = "DataFile/demo.xlsx";

        //'fileinputstream' class to open the file (import from io. and it throws some exceptions)
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(excelFilePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        //get workbook
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }

        //get sheets from workbook
        XSSFSheet sheet = workbook.getSheet("Sheet1"); //Bring workbook in scope


        //--------------------------------------------iterator-------------------------------------------
        Iterator iterator = sheet.iterator(); //add sheet

        while (iterator.hasNext()) {

            //iterator for ROW and get it
            XSSFRow row = (XSSFRow) iterator.next(); //type casting  , it has row which has cells

            Iterator cellIterator = row.cellIterator(); //iterate all the cells in this row

            //iterator for cell and get it

            while (cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();
                //----------------------------------------------------------------------------------
                //to verify and read the data we use Switch statement
                //get values by different data types
                switch (cell.getCellType()) {  //bring cell into the scope
                    //get string value
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    //get numeric value
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                    //get boolean condition
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                }
                System.out.print(" | ");
            }
            System.out.println();
        }
    }
}
