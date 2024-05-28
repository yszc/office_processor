package com.laotie.app;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

import org.junit.jupiter.api.Test;

import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;

class SectionTest {
    String filePath = "/Users/chenzhijun/workbench/wordparser/docs/formDefine.json";


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
    
    /**
     * 降维
     * @throws IOException
     */
    @Test
    void test1D() throws IOException {
        String formJson = readFileAsString("docs/formDefine.json");
        Section root = Section.fromJson(formJson);
        List<JSONObject> formList = root.fetchAllInputAttr(false);
        // filter the item which has no attribute var_name
        formList = formList.stream()
        .filter(item -> item.getString("position_title")==null || item.getString("position_title").isEmpty())
        .collect(Collectors.toList());
        System.out.println(formList);
    }


}