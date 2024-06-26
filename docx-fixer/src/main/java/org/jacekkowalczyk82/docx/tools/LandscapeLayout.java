/*
 * This Java source file was generated by the Gradle 'init' task.
 */
package org.jacekkowalczyk82.docx.tools;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;


public class LandscapeLayout {


    public static void main(String[] args) {

        String inputFileName = null;
        String format = null;

        if (args.length == 0) {
            usage();
            System.exit(1);
        } else if (args.length == 1) {
            inputFileName = args[0];
            format = "A4";
        } else if (args.length == 2) {
            inputFileName = args[0];
            format = args[1];
        }

        if (inputFileName != null && format != null) {
            System.out.println("Setting Landscape page layout");

            try {


                XWPFDocument document = new XWPFDocument(new FileInputStream(inputFileName));

                // Set page layout to landscape
                CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
                XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);
                CTPageSz pageSize = sectPr.addNewPgSz();
                pageSize.setOrient(STPageOrientation.LANDSCAPE);

                //https://stackoverflow.com/questions/20188953/how-to-set-page-orientation-for-word-document

                // A4
                if (format.equals("A4")) {
                    //A4           595x842
                    pageSize.setW(BigInteger.valueOf(842 * 20)); // 11.69 inch (A4 landscape width)
                    pageSize.setH(BigInteger.valueOf(595 * 20)); // 8.27 inch (A4 landscape height)
                } else if (format.equals("A3")) {
                    //A3           842x1190
                    pageSize.setW(BigInteger.valueOf(1190 * 20)); //
                    pageSize.setH(BigInteger.valueOf(842 * 20)); //
                } else {
                    throw new UnsupportedOperationException("This "+ format + " is not supported yet");
                }

                // A3


                FileOutputStream fileOut = new FileOutputStream(inputFileName + "_" + format+ "_landscape.docx");
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
        System.out.println("Setting Landscape page layout");
        System.out.println("Required parameters were not provided");
        System.out.println("Usage: landscapelayout.bat <PATH TO INPUT DOCX FILE>");
    }
}
