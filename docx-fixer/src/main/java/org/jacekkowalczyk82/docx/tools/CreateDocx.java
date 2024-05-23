package org.jacekkowalczyk82.docx.tools;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class CreateDocx {

    public static void main(String[] args) throws IOException {
        XWPFDocument document = new XWPFDocument();

        // Create a header
        XWPFParagraph header = document.createParagraph();
        XWPFRun run = header.createRun();
        run.setBold(true);
        run.setFontSize(20);
        run.setText("Header");

        // Create a paragraph
        XWPFParagraph paragraph = document.createParagraph();
        run = paragraph.createRun();
        run.setText("This is a sample paragraph.");

        // Create a table
        XWPFTable table = document.createTable();
        XWPFTableRow row = table.createRow();
        row.createCell().setText("Column 1");
        row.createCell().setText("Column 2");
        row = table.createRow();
        row.createCell().setText("Row 2, Col 1");
        row.createCell().setText("Row 2, Col 2");

        // Create a bullet list
        XWPFParagraph bulletList = document.createParagraph();
        bulletList.setIndentationLeft(720);
        XWPFRun bulletRun = bulletList.createRun();
        bulletRun.setText("Bullet Point 1");
        bulletList = document.createParagraph();
        bulletList.setIndentationLeft(720);
        bulletRun = bulletList.createRun();
        bulletRun.setText("Bullet Point 2");

        try (FileOutputStream out = new FileOutputStream("output.docx")) {
            document.write(out);
        }
        document.close();
    }
}