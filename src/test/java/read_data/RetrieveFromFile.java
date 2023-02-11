package read_data;

import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class RetrieveFromFile {
    @Test
    public void readFileTest() throws IOException {
        File excelFile = new File("src/test/resources/Book1.xlsx");
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet page1 = workbook.getSheet("Sheet1");
        XSSFRow row1 = page1.getRow(0);
        XSSFCell cell1 = row1.getCell(0);
        System.out.println(cell1);
        XSSFRow row2 = page1.getRow(1);
        XSSFCell cell2 = row2.getCell(0);
        System.out.println(cell2);

    }

    @Test
    public void getRowValuesTest() throws IOException {
        File file = new File("src/test/resources/Book1.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0);
        XSSFRow row1 = sheet1.getRow(0);

        for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum(); i++) {
            XSSFCell cell = row1.getCell(i);
            System.out.print(cell + " | ");
        }


        }
        @Test
        public void getRowValuesTest1 () throws IOException {
             File file1 = new File("src/test/resources/Book1.xlsx");
            FileInputStream fileInputStream1 = new FileInputStream(file1);
            XSSFWorkbook workbook1 = new XSSFWorkbook(fileInputStream1);
            XSSFSheet sheet1= workbook1.getSheetAt(0);


            for(int i= sheet1.getFirstRowNum();i< sheet1.getLastRowNum();i++){
               XSSFRow temp=  sheet1.getRow(i);
                System.out.print(" | ");
               for(int j= temp.getFirstCellNum();j< temp.getLastCellNum();j++){
                   System.out.print(temp.getCell(j)+ " | ");
               }

                System.out.println();
            }

        }
        @Test
    public void getAllData() throws IOException {
            File file1 = new File("src/test/resources/TestData.xlsx");
            FileInputStream fileInputStream1 = new FileInputStream(file1);
            XSSFWorkbook workbook1 = new XSSFWorkbook(fileInputStream1);
            XSSFSheet sheet1= workbook1.getSheetAt(0);
            for(int i= sheet1.getFirstRowNum();i< sheet1.getLastRowNum();i++){
                XSSFRow temp=  sheet1.getRow(i);
                System.out.print(" | ");
                for(int j= temp.getFirstCellNum();j< temp.getLastCellNum();j++){
                    System.out.print(temp.getCell(j)+ " | ");
                }

                System.out.println();
            }


        }
}