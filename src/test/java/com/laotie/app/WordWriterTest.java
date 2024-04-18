package com.laotie.app;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;

import com.alibaba.fastjson2.JSON;
import org.junit.jupiter.api.Test;

public class WordWriterTest {


    @Test
    public void testWriteTemplate() throws IOException {
        String formData = readFileAsString("docs/formData.json");
        try {
            WordWriter writer = new WordWriter("docs/template.docx", JSON.parseObject(formData));
            writer.writeTemplate("docs/output.docx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String readFileAsString(String filePath) throws IOException {
        StringBuilder fileContent = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new FileReader(filePath))) {
            String line;
            while ((line = reader.readLine()) != null) {
                fileContent.append(line).append("\n");
            }
        }
        return fileContent.toString().trim(); // Remove trailing newline
    }
}
