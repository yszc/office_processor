package com.laotie.app;
import java.io.IOException;

import org.junit.jupiter.api.Test;

class WordParserTest {

    @Test
    void testParseTemplate() throws IOException {
        String filePath = "docs/jsoncase.docx";
        WordParser wordParser = new WordParser(filePath);
        Section root = wordParser.parseTemplate();

        String jsonResult = root.toFormFriendly().toNoneEmpty().toJson();
        System.out.println(jsonResult);
    }

}