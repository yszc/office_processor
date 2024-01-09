package com.laotie.app;

import org.apache.poi.xwpf.usermodel.*;
import java.io.FileOutputStream;

public class Writer {
    public static void main(String[] args) throws Exception {
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("This is a paragraph. ");

        // Create a table
        XWPFTable table = document.createTable(3, 3); // 3 rows, 3 columns

        // Set values to cells
        for (int row = 0; row < 3; row++) {
            for (int col = 0; col < 3; col++) {
                table.getRow(row).getCell(col).setText("row " + row + ", col " + col);
            }
        }

        run = paragraph.createRun();
        run.addBreak();
        run.setText("This is another paragraph.");

        FileOutputStream out = new FileOutputStream("output.docx");
        document.write(out);
        out.close();
        document.close();
    }
}
