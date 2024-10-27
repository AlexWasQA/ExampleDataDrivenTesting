import org.apache.commons.csv.*;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import static org.junit.jupiter.api.Assertions.assertEquals;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class CSV {
    @Test
    @Order(1)
    public void readCSV() throws IOException {
        // read and print all information from csv file
        String csvFilePath = "data.csv";
        Reader reader = new FileReader(csvFilePath);
        CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT);
        for (CSVRecord csvRecord : csvParser) {
            for (String field : csvRecord) {
                System.out.print(field + " ");
            }
            System.out.println();
        }
        csvParser.close();
        reader.close();
    }

    @Test
    @Order(2)
    public void readLine6() throws IOException {
        String csvFilePath = "data.csv";
        Reader reader = new FileReader(csvFilePath);
        CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT);
        int targetLine = 6; // Line number to read (1-based)
        int currentLine = 0;
        for (CSVRecord csvRecord : csvParser) {
            if (currentLine == targetLine) {
                for (String field : csvRecord) {
                    System.out.print(field + " ");
                }
                System.out.println();
                break; // Stop processing after reading the target line
            }
            currentLine++;
        }
        csvParser.close();
        reader.close();
    }

    @Test
    @Order(3)
    public void findFilmsAfter2007() throws IOException {
        String csvFilePath = "data.csv";
        Reader reader = new FileReader(csvFilePath);
        CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT
                .withFirstRecordAsHeader() // Assuming the first row is a header
                .withTrim());
        csvParser.iterator().next();
        for (CSVRecord csvRecord : csvParser) {
            String yearStr = csvRecord.get("Year");
            try {
                int year = Integer.parseInt(csvRecord.get("Year"));
                if (year > 2007) {
                    System.out.println(csvRecord.get("Title") + " (" + csvRecord.get("Year") + ")");
                }
            } catch (NumberFormatException e) {

            }
        }
        System.out.println();
        csvParser.close();
        reader.close();
    }

    @Test
    @Order(4)
    public void addDataToCSV() throws IOException {
        String csvFilePath = "data.csv";
        int lineCountBefore = countLines(csvFilePath);
        String newData = "15, Inside Out, Pete Docter, 2015, 95";
        FileWriter fileWriter = new FileWriter(csvFilePath, true);
        BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);
        bufferedWriter.newLine();
        bufferedWriter.write(newData);
        bufferedWriter.close();
        // Reset the reader's position
        int lineCountAfter = countLines(csvFilePath);
        assertEquals(lineCountBefore +1, lineCountAfter);
    }

    private int countLines(String filePath) throws IOException {
        BufferedReader reader = new BufferedReader(new FileReader(filePath));
        int lineCount = 0;
        while (reader.readLine() != null) {
            lineCount++;
        }
        reader.close();
        return lineCount;
    }

    @Test
    @Order(5)
    public void deleteInformationFromCSV() throws IOException {
        // delete last line
        String csvFilePath = "data.csv";
        String csvCopyPath = "data_copy.csv";
        copyFile(csvFilePath, csvCopyPath);
        // Perform the operation to remove the last line
        List<String> modifiedContent = removeLastLineFromCSV(csvFilePath);
        // Read the content of the original CSV copy
        List<String> originalContent = readCSVContent(csvCopyPath);
        // Verify that the modified content matches the original content minus the last line
        assertEquals(originalContent.subList(0, originalContent.size() - 1), modifiedContent);
    }
        private void copyFile(String sourcePath, String destinationPath) throws IOException {
            try (InputStream in = new FileInputStream(sourcePath);
                 OutputStream out = new FileOutputStream(destinationPath)) {
                byte[] buffer = new byte[1024];
                int length;
                while ((length = in.read(buffer)) > 0) {
                    out.write(buffer, 0, length);
                }
            }
        }
    private List<String> removeLastLineFromCSV(String filePath) throws IOException {
        List<String> lines = new ArrayList<>();
        BufferedReader reader = new BufferedReader(new FileReader(filePath));
        String line;
        while ((line = reader.readLine()) != null) {
            lines.add(line);
        }
        reader.close();

        if (!lines.isEmpty()) {
            lines.remove(lines.size() - 1);
        }
        BufferedWriter writer = new BufferedWriter(new FileWriter(filePath));
        for (String modifiedLine : lines) {
            writer.write(modifiedLine);
            writer.newLine();
        }
        writer.close();
        return lines;
    }
    private List<String> readCSVContent(String filePath) throws IOException {
        List<String> content = new ArrayList<>();
        BufferedReader reader = new BufferedReader(new FileReader(filePath));
        String line;
        while ((line = reader.readLine()) != null) {
            content.add(line);
        }
        reader.close();
        return content;
    }
}


