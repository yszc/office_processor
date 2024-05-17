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
    @Test
    void test1D() throws IOException {
        String formJson = readFileAsString("docs/constant.json");
        Section root = Section.fromJson(formJson);
        List<JSONObject> formList = root.fetchAllInputAttr(false);
        // filter the item which has no attribute var_name
        formList = formList.stream()
        .filter(item -> !(item.containsKey("var_name") && !item.getString("var_name").isEmpty()) )
        .collect(Collectors.toList());
        System.out.println(formList);
    }


}