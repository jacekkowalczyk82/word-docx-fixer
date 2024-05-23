package org.jacekkowalczyk82.docx.tools;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;


public class CreateLandscapeDocx {
    public static void main(String[] args) {
        XWPFDocument document = new XWPFDocument();

        // Set page layout to landscape
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);
        CTPageSz pageSize = sectPr.addNewPgSz();
        pageSize.setOrient(STPageOrientation.LANDSCAPE);
        pageSize.setW(BigInteger.valueOf(16840)); // 11.69 inch (A4 landscape width)
        pageSize.setH(BigInteger.valueOf(11900)); // 8.27 inch (A4 landscape height)

        // Add a header
        XWPFParagraph headerParagraph = document.createParagraph();
        XWPFRun headerRun = headerParagraph.createRun();
        headerRun.setText("This is the Header");
        headerRun.setBold(true);

        // Add a title
        XWPFParagraph titleParagraph = document.createParagraph();
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText("Document Title");
        titleRun.setBold(true);
        titleRun.setFontSize(20);

        // Add a paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("This is a sample paragraph in the document.");

        // Add a bullet list
        XWPFParagraph bulletParagraph1 = document.createParagraph();
        bulletParagraph1.setStyle("ListBullet");
        XWPFRun bulletRun1 = bulletParagraph1.createRun();
        bulletRun1.setText("Bullet Item 1");

        XWPFParagraph bulletParagraph2 = document.createParagraph();
        bulletParagraph2.setStyle("ListBullet");
        XWPFRun bulletRun2 = bulletParagraph2.createRun();
        bulletRun2.setText("Bullet Item 2");

        // Add a table
        XWPFTable table = document.createTable(3, 3);
        for (int row = 0; row < 3; row++) {
            for (int col = 0; col < 3; col++) {
                table.getRow(row).getCell(col).setText("Cell " + (row + 1) + "," + (col + 1));
            }
        }

        // Save the document
        try (FileOutputStream out = new FileOutputStream("generated-doc.docx")) {
            document.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Document created successfully!");
    }
}


