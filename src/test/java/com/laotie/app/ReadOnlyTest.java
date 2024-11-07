package com.laotie.app;

import org.apache.poi.xwpf.usermodel.*;
import java.io.FileOutputStream;
import java.io.IOException;

public class ReadOnlyTest {
    public static void main(String[] args) throws IOException {
        String fileName = "hello.docx";

        try (XWPFDocument doc = new XWPFDocument()) {
            doc.enforceReadonlyProtection();
            // Create a paragraph
            XWPFParagraph p1 = doc.createParagraph();
            p1.setAlignment(ParagraphAlignment.CENTER);

            // Set font
            XWPFRun r1 = p1.createRun();
            r1.setBold(true);
            r1.setItalic(true);
            r1.setFontSize(22);
            r1.setFontFamily("New Roman");
            r1.setText("I am the first paragraph.");

            // Save it to .docx file
            try (FileOutputStream out = new FileOutputStream(fileName)) {
                doc.write(out);
            }
        }
    }
}
