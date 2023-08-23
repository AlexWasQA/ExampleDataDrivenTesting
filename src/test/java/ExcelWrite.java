import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWrite {
        String filePath = "Movies.xlsx";
        XSSFWorkbook workbook;
        XSSFSheet sheet;
        FileInputStream fileInputStream;
        {
                try {
                        fileInputStream = new FileInputStream(filePath);
                } catch (FileNotFoundException e) {
                        throw new RuntimeException(e);
                }
        }
        @Before
        public void setUp() throws IOException {
                workbook = new XSSFWorkbook(fileInputStream);
                sheet = workbook.getSheet("Movies");
        }
        @After
        public void tearDown() throws IOException{
                FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                workbook.write(fileOutputStream);
                fileOutputStream.close();
                workbook.close();
                fileInputStream.close();
        }
        @Test
        public void excelCreateNewColumn(){
                XSSFCell ratingRow = sheet.getRow(0).createCell(5);
                ratingRow.setCellValue("Rating");
        }
        @Test
        public void excelWriteValueInNewColumn(){
                int lastMovieRow = sheet.getLastRowNum();
                //set values for first 2 rows
                XSSFCell rating1 = sheet.getRow(1).createCell(5);
                rating1.setCellValue(8.3);
                XSSFCell rating2 = sheet.getRow(2).createCell(5);
                rating2.setCellValue(7.2);
                //set value for last row
                XSSFCell ratingLast = sheet.getRow(lastMovieRow).createCell(5);
                ratingLast.setCellValue(7.4);
                /**
                 *  set value for middle row. Excel has row 0 with names of columns. It is mean I should
                 *  exclude row 0 for correct calculation from row's count.
                 */
                int rowCount = sheet.getPhysicalNumberOfRows();
                if(rowCount%2 == 0){
                        int middleRow = rowCount / 2;
                        XSSFCell middleRowValue = sheet.getRow(middleRow).createCell(5);
                        middleRowValue.setCellValue(7.2);
                }else{
                        int middleRow1 = lastMovieRow / 2;
                        XSSFCell middleRowValue1 = sheet.getRow(middleRow1).createCell(5);
                        middleRowValue1.setCellValue(7.2);
                        int middleRow2 = lastMovieRow / 2 + 1;
                        XSSFCell middleRowValue2 = sheet.getRow(middleRow2).createCell(5);
                        middleRowValue2.setCellValue(8);
                }
        }
        @Test
        public void excelWriteNewRow(){
                int lastMovieRow = sheet.getLastRowNum();
                XSSFRow newMovie = sheet.createRow(lastMovieRow);
                newMovie.createCell(0).setCellValue(15);
                newMovie.createCell(1).setCellValue("Inside Out");
                newMovie.createCell(2).setCellValue("Pete Docter");
                newMovie.createCell(3).setCellValue(2015);
                newMovie.createCell(4).setCellValue(95);

        }
        @Test
        public void excelDeleteRow(){
              int lastMovieRow = sheet.getPhysicalNumberOfRows();
              int rowNum = 5;
              XSSFRow rowForDelete = sheet.getRow(rowNum);
              sheet.removeRow(rowForDelete);
        }
        @Test
        public void excelDeleteColumn() {
                int colNumToRemove = 4;
                for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                        XSSFRow row = sheet.getRow(i);
                        for (int j = 0; j < row.getLastCellNum(); j++) {
                                if (j == colNumToRemove) {
                                        row.removeCell(row.getCell(j));
                                }
                        }
                }
        }
        @Test
        public void excelDeleteAllDataExcludesColumnNames () {
                                XSSFRow row;
                                for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                                        row = sheet.getRow(i);
                                        for (int j = 0; j < row.getLastCellNum(); j++) {
                                                row.removeCell(row.getCell(j));
                                        }
                                }
                        }
}
