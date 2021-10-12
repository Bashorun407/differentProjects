package ExcelIterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class IteratorReader {
    public static void main(String[] args) throws IOException {

        String filePath ="C:\\Users\\Akinbobola Oluwaseyi\\Desktop\\TestingFile\\src\\main\\java\\ExcelData\\Data1.xlsx";

        //Reading the file using FILEINPUT STREAM
        FileInputStream inputStream = new FileInputStream(filePath);

        //Getting the workbook from the excel file
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        //Getting the specific sheet from the workbook
        XSSFSheet sheet = workbook.getSheet("Sheet2");

        //Getting the rows from the specified sheet
        int rows = sheet.getLastRowNum();

        //Getting the number of columns in each row
        int columns = sheet.getRow(1).getLastCellNum();

        //Using 2 FOR LOOPS to read and print the contents of each row and columns
        for(int i = 0; i<=rows; i++){

            XSSFRow row = sheet.getRow(i);
            for(int j = 0; j<columns; j++){

                XSSFCell cell = row.getCell(j);

                //For MULTIPLE DATA TYPE
                switch (cell.getCellType()){

                    //If the datatype is String
                    case STRING -> {
                        System.out.println(cell.getNumericCellValue() + " |");
                        break;
                    }

                    //If the datatype is Numeric
                    case NUMERIC -> {
                        System.out.println(cell.getNumericCellValue() + "  |");
                        break;
                    }

                    //If the datatype is Boolean
                    case BOOLEAN -> {
                        System.out.println(cell.getBooleanCellValue() + "  |");
                        break;
                    }
                }

            }
            //NEXT LINE
            System.out.println();
        }

        //USING THE WHILE LOOP TO READ DATA FROM THE SPECIFIED EXCEL FILE

        Iterator iterator = sheet.iterator();
        while (iterator.hasNext()){
            XSSFRow row = (XSSFRow) iterator.next();

            Iterator cellIterator = row.cellIterator();

            while(cellIterator.hasNext()){

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

                //TO PRINT DATA ON A NEW LINE
                System.out.println();
            }
        }
    }
}
