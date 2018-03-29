package poiExample;


import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExcelReader {

    public static void main(String[] args) throws IOException {
        InputStream inputStream = new FileInputStream("src\\poiExample\\excel\\users.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell name = row.getCell(0);
            XSSFCell surname = row.getCell(1);
            XSSFCell email = row.getCell(2);
            XSSFCell password = row.getCell(3);
            XSSFCell gender = row.getCell(4);

            System.out.print(name.getStringCellValue() + " \t");
            System.out.print(surname.getStringCellValue() +" \t");
            System.out.print(email.getStringCellValue() + " \t");
         //   if(password.getCellTypeEnum() == CellType.NUMERIC){
                System.out.print(formatter.formatCellValue(password) + " \t");
           // }else{
             //   System.out.print(password.getStringCellValue() + " \t");
           // }
            System.out.print(gender.getStringCellValue() + " \t");
            System.out.println();
        }
    }

}
