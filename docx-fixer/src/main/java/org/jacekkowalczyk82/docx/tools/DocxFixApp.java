package org.jacekkowalczyk82.docx.tools;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;


public class DocxFixApp {


    public static void main(String[] args) {

        String inputFileName = null;
        String options = null;
        String format = null;

        if (args.length == 0) {
            usage();
            System.exit(1);
        } else if (args.length == 1) {
            inputFileName = args[0];
            options = "LANDSCAPE_A4_SOLID_TABLE_BORDERS";
            format = "A4";
        } else if (args.length == 2) {
            inputFileName = args[0];
            options = args[1];
        }

        if (inputFileName != null && options != null) {
            System.out.println("Docx-fixer");

            try {


                XWPFDocument document = new XWPFDocument(new FileInputStream(inputFileName));

                if (options.contains("LANDSCAPE")) {
                    // Set page layout to landscape

                    //https://stackoverflow.com/questions/20188953/how-to-set-page-orientation-for-word-document
                    CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
                    XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);
                    CTPageSz pageSize = sectPr.addNewPgSz();
                    pageSize.setOrient(STPageOrientation.LANDSCAPE);

                    // A4
                    if (options.contains("A4")) {
                        //A4           595x842
                        pageSize.setW(BigInteger.valueOf(842 * 20)); // 11.69 inch (A4 landscape width)
                        pageSize.setH(BigInteger.valueOf(595 * 20)); // 8.27 inch (A4 landscape height)
                    } else if (options.contains("A3")) {
                        //A3           842x1190
                        pageSize.setW(BigInteger.valueOf(1190 * 20)); //
                        pageSize.setH(BigInteger.valueOf(842 * 20)); //
                    } else {
                        throw new UnsupportedOperationException("This "+ options + " is not supported yet");
                    }
                }

                if (options.contains("SOLID_TABLE_BORDERS")) {
                    for (XWPFTable table : document.getTables()) {
                        table.setWidth("100%");

                        //set borders
                        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
                        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
                        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
                        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
                        table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");
                        table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000");

                    }
                }

                FileOutputStream fileOut = new FileOutputStream(inputFileName + "_" + options+ ".docx");
                document.write(fileOut);
                fileOut.close();
                document.close();

                System.out.println("DONE");

            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        }

    }

    public static void usage() {
        System.out.println("The application can fix some settings of the docx file like page laout and solid table borders");
        System.out.println("Required parameters were not provided");
        System.out.println("Usage: docx-fixer.bat <PATH TO INPUT DOCX FILE> <options>");
        System.out.println("    Example: docx-fixer.bat example.docx LANDSCAPE_A3_SOLID_TABLE_BORDERS");
    }
}
