package ExcelFileReader;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {
    public static void main(String[] args) throws IOException {


        String excelFilePath = "C:\\Users\\Akinbobola Oluwaseyi\\Desktop\\TestingFile\\src\\main\\java\\ExcelData\\Data1.xlsx";

        FileInputStream inputStream = new FileInputStream(excelFilePath);

        //To get the workbook in the excel file
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        //To get the name of the sheet to work with
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        //To get the number of rows in the sheet
        int rows = sheet.getLastRowNum();

        int columns = sheet.getRow(1).getLastCellNum();

        //using FOR LOOP to iterate through the cells of a row before moving to the next row

       // System.out.println("\n\n\n PRINTING THE EXCEL FILE USING FOR LOOP");
//        for(int i= 0; i<=rows; i++){
//
//            XSSFRow row = sheet.getRow(i);
//
//            for(int j = 0; j<columns; j++){
//
//                XSSFCell cell = row.getCell(j);
//
//                //for MULTIPLE DATA TYPE, USE SWITCH STATEMENT
//                switch (cell.getCellType()){
//                    case STRING ->{
//                        System.out.print(cell.getStringCellValue() + "    |");
//                        break;
//                    }
//                    case NUMERIC -> {
//                        System.out.print(cell.getNumericCellValue() + "    |");
//                        break;
//                    }
//
//                    case BOOLEAN -> {
//                        System.out.print(cell.getBooleanCellValue() + "    |");
//                        break;
//                    }
//
//                }
//
//            }
//
//            System.out.println();
//        }


        //CREATING A DEMARCATION TO SEPARATE THE FOR LOOP FROM THE WHILE LOOP
        System.out.println("\n\n\n THE FOLLOWING IS READ USING THE FOR LOOP!!!");

        //READING DATA FROM THE EXCEL SHEET USING THE ITERATOR...code is easier to write and understand
        //CREATING AN ITERATOR VARIABLE/CLASS/OBJECT
        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()) {
            XSSFRow row = (XSSFRow) iterator.next();

            Iterator cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();

                //for MULTIPLE DATA TYPE, USE SWITCH STATEMENT
                switch (cell.getCellType()){
                    case STRING ->{
                        System.out.print(cell.getStringCellValue() + "    |");
                        break;
                    }
                    case NUMERIC -> {
                        System.out.print(cell.getNumericCellValue() + "    |");
                        break;
                    }

                    case BOOLEAN -> {
                        System.out.print(cell.getBooleanCellValue() + "    |");
                        break;
                    }

                }

            }

            //TO PRINT ON A NEW LINE
            System.out.println();
        }


    }
}
