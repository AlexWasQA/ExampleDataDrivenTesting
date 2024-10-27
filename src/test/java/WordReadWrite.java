import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;


public class WordReadWrite {
    @Test
    public void readOriginalWordFile() throws IOException {
        // find first paragraph
        XWPFDocument document = new XWPFDocument(new FileInputStream("lorem_ipsum.docx"));
        String firstPar = document.getParagraphs().get(0).getText();
        // count of paragraphs
        /**
         * The count of paragraphs in a Word file includes the blank paragraph
         * at the end of the document. This blank paragraph is not visible, but it is still there.
         */
        int countP = document.getParagraphs().size();
        // find last paragraph.
        String lastPar = document.getParagraphs().get(countP-2).getText();
    }
    @Test
    public void removeBlankParagraphWithCreatingNewFile() throws IOException{
        /**
         * The count of paragraphs in a Word file includes the blank paragraph at the end of the document.
         * This blank paragraph is not visible, but it is still there.
          */
        XWPFDocument document = new XWPFDocument(new FileInputStream("lorem_ipsum.docx"));
        int countP = document.getParagraphs().size();
        // Iterate through paragraphs and find blank ones
        List<XWPFParagraph> blankParagraphs = new ArrayList<>();
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            if (paragraph.getText().trim().isEmpty()) {
                blankParagraphs.add(paragraph);
            }
        }
        // Remove blank paragraphs
        for(XWPFParagraph paragraph : blankParagraphs){
            document.removeBodyElement(document.getPosOfParagraph(paragraph));
        }
        FileOutputStream outputStream = new FileOutputStream("lorem_ipsum_without_blank_par.docx");
        document.write(outputStream);
        XWPFDocument document1 = new XWPFDocument(new FileInputStream("lorem_ipsum_without_blank_par.docx"));
        int countNew = document1.getParagraphs().size();
        // creating assertions for check different numbers of paragraphs original file and new
        assertNotEquals(countNew, countP);
        outputStream.close();
        /**
         * Also I can remove blank paragraph from Word file without creating new file:
         * For that I find blank paragraphs like in previous method and remove its using next script:
         *          for (XWPFParagraph paragraph : paragraphsToRemove) {
         *             int pos = document.getPosOfParagraph(paragraph);
         *             if (pos >= 0) {
         *                 document.removeBodyElement(pos);
         *             }
         *         FileOutputStream outputStream = new FileOutputStream("lorem_ipsum.docx");
         *         document.write(outputStream);
         */
    }
    @Test
    public void addInformationInNewFile() throws IOException{
        //creating new paragraph
        XWPFDocument document = new XWPFDocument(new FileInputStream("lorem_ipsum_without_blank_par.docx"));
        int countOrigin = document.getParagraphs().size();
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.createRun().setText("It is new paragraph with new information.");
        FileOutputStream outputStream = new FileOutputStream("lorem_ipsum_without_blank_par.docx");
        document.write(outputStream);
        outputStream.close();
        /** for checking success creation I will use assertion of numbers of paragraphs (number should be greater
         *  than previous) and assertion of containing text in file
          */
        int countAfterAdd = document.getParagraphs().size();
        assertNotEquals(countOrigin, countAfterAdd);
        String newPar = "It is new paragraph with new information.";
        assertTrue(document.getParagraphs().get(countAfterAdd-1).getText().equals(newPar));
    }
    @Test
    public void checkInformationInNewFile() throws IOException{
        // find all spaces in text
        XWPFDocument document = new XWPFDocument(new FileInputStream("lorem_ipsum_without_blank_par.docx"));
        int spaceCount = 0;
        // Iterate through paragraphs and runs to find spaces
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0); // Get the text content of the run
                if (text != null) {
                    for (int i = 0; i < text.length(); i++) {
                        if (Character.isWhitespace(text.charAt(i))) {
                            spaceCount++;
                        }
                    }
                }
            }
        }
        // find word 'massa' and how many times it is present in text
        String word = "massa";
        int wordCount = 0;
        for (XWPFParagraph paragraph : document.getParagraphs()){
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0); // Get the text content of the run
                if (text != null && text.contains(word)) {
                    // Count occurrences of the target word in the current run
                    int occurrences = (text.length() - text.replace(word, "").length()) / word.length();
                    wordCount += occurrences;
                }
            }
        }
        // Close resources
        document.close();
    }
    @Test
    public void deletePartOfInformation() throws IOException {
        //I will delete previous than last paragraph and create new file
        XWPFDocument document = new XWPFDocument(new FileInputStream("lorem_ipsum_without_blank_par.docx"));
        int countP = document.getParagraphs().size();
        document.removeBodyElement(countP - 2);
        document.write(new FileOutputStream("new.docx"));
    }
}
