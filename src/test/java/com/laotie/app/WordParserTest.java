package com.laotie.app;
import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;

class WordParserTest {
    String filePath = "docs/template.docx";

    @Test
    void testParseTemplate() throws IOException {
        WordParser wordParser = new WordParser(filePath);
        Section root = wordParser.parseTemplate();

        String jsonResult = root.toFormFriendly().toNoneEmpty().toJson();
        System.out.println(jsonResult);
    }

    @Test
    void testFormList() throws IOException {
        WordParser wordParser = new WordParser(filePath);
        Section root = wordParser.parseTemplate();
        List<JSONObject> formList = root.fetchAllInputAttr();
        System.out.println(new JSONArray(formList));
    }


}