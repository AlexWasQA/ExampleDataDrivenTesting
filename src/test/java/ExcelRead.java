import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import static org.junit.Assert.assertTrue;


public class ExcelRead {
    @Test
    public void read() throws IOException {
        File file = new File("Movies.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Movies");
        assertTrue(sheet.getRow(0).getCell(0).toString().equals("Id"));
        assertTrue(sheet.getRow(4).getCell(2).toString().equals("Pete Docter"));
    }

    @Test
    public void findJohnLasseterMovies() throws IOException {
        File file = new File("Movies.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Movies");
        int usesRows = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < usesRows; rowNum++) {
            if (sheet.getRow(rowNum).getCell(2).toString().equals("John Lasseter")) {
                ArrayList<String> movies = new ArrayList<>();
                movies.add(String.valueOf(sheet.getRow(rowNum).getCell(1)));
                String str = String.join(",", movies);
                System.out.println(str);
            }
        }
    }

    @Test
    public void findAllInformationAboutJohnLasseterMovies() throws IOException {
        File file = new File("Movies.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Movies");
        int usesRows = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < usesRows; rowNum++) {
            if (sheet.getRow(rowNum).getCell(2).toString().equals("John Lasseter")) {
                ArrayList<String> movies = new ArrayList<>();
                for (int colNum = 1; colNum < sheet.getRow(rowNum).getLastCellNum(); colNum++) {
                    movies.add(String.valueOf(sheet.getRow(rowNum).getCell(colNum)));
                }
                String str = String.join(",", movies);
                System.out.println(str);
            }
        }
    }

    @Test
    public void findMoviesAfter2007() throws IOException {
        File file = new File("Movies.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Movies");
        int usesRows = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < usesRows; rowNum++) {
            // Check if the year of the movie is after 2007
            double year = Double.parseDouble(sheet.getRow(rowNum).getCell(3).toString());
            if (year > 2007) {
                ArrayList<String> movies = new ArrayList<>();
                for (int colNum = 1; colNum < sheet.getRow(rowNum).getLastCellNum(); colNum++) {
                    movies.add(String.valueOf(sheet.getRow(rowNum).getCell(colNum)));
                }
                String str = String.join(",", movies);
                System.out.println(str);
            }
        }
    }

    @Test
    public void findAllInformationWithID() throws IOException {
        File file = new File("Movies.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        // Read the data from the first sheet
        XSSFSheet sheet1 = workbook.getSheet("Movies");
        int lastCellNum = sheet1.getLastRowNum();
        ArrayList<String> information1 = new ArrayList<>();
        for (int i = 0; i < lastCellNum; i++) {
            Cell cell = sheet1.getRow(1).getCell(i);
            if (cell != null) {
                information1.add(cell.toString());

            }
        }
        // Read the data from the second sheet
        XSSFSheet sheet2 = workbook.getSheet("Rating");
        int lastCellNum2 = sheet2.getLastRowNum();
        ArrayList<String> information2 = new ArrayList<>();
        for (int i = 0; i < lastCellNum2; i++) {
            Cell cell = sheet2.getRow(1).getCell(i);
            if (cell != null) {
                information2.add(cell.toString());
            }
        }
        // Print the information from the two sheets
        System.out.println("Information from both sheets: " + information1 + information2);
    }
}




