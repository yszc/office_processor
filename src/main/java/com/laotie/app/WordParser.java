package com.laotie.app;

import org.apache.poi.xwpf.usermodel.*;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONException;
import com.alibaba.fastjson2.JSONObject;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;

public class WordParser {
    protected XWPFDocument document;
    protected XWPFStyles style_sheet;
    protected Section nestedRoot;

    public WordParser(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath)) {
            this.initWordParser(fis);
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        }
    }

    public WordParser(InputStream fis){
        try {
            this.initWordParser(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private void initWordParser(InputStream fis) throws IOException {
        this.document = new XWPFDocument(fis);
        // 样式表
        this.style_sheet = document.getStyles();
        this.nestedRoot = new Section("root", "root");
    }

    /**
     * 解析模板，获得标题和对应的输入框信息
     * 
     * @param document
     * @return
     */
    public Section parseTemplate() {
        Stack<Section> stack = new Stack<>();
        this.nestedRoot = new Section("root", "root");
        stack.push(this.nestedRoot);

        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                parseParagraph(paragraph, stack);
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                parseTable(table, stack);
            }
        }
        return nestedRoot;
    }

    /**
     * 解析table，从每个单元格中解析占位符。
     * 
     * @param table
     * @param stack
     */
    private void parseTable(XWPFTable table, Stack<Section> stack) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (Section input : parseInput(cell.getText())) {
                    stack.peek().getChildren().add(input);
                }
            }
        }
    }

    /**
     * 解析段落信息，包括heading和普通段落
     * 
     * @param paragraph
     * @param stack
     */
    private void parseParagraph(XWPFParagraph paragraph, Stack<Section> stack) {
        String paraText = paragraph.getText();
        String styleID = paragraph.getStyleID();
        String styleName = "";
        if (styleID != null) {
            styleName = style_sheet.getStyle(styleID).getName().toLowerCase();
        }
        if (styleName.startsWith("heading")) {
            int indent = parseHeadingIndent(styleName);
            if (indent <= 0) {
                return;
            }
            while (stack.size() > indent) {
                stack.pop();
            }
            Section paraSection = new Section("title", paraText);
            if (!stack.isEmpty()) {
                stack.peek().getChildren().add(paraSection);
            }
            String parentPrefix = stack.peek().getPrefix();
            List<Section> brothers = stack.peek().getChildren();
            if (parentPrefix.length() > 0) {
                paraSection.setPrefix(parentPrefix + "."
                        + String.valueOf( brothers.stream().filter(section -> "title".equals(section.getType())).count()));
            } else {
                paraSection.setPrefix(
                        String.valueOf(brothers.stream().filter(section -> "title".equals(section.getType())).count()));
            }
            stack.push(paraSection);
        } else {
            // 查找占位符
            for (Section input : parseInput(paraText)) {
                stack.peek().getChildren().add(input);
            }
        }
    }

    /**
     * 获得标题的层级
     * 
     * @param heading
     * @return
     */
    private static int parseHeadingIndent(String heading) {
        if (heading.startsWith("heading")) {
            return Integer.parseInt(heading.replace("heading", "").trim());
        }
        return 0;
    }

    /**
     * 解析输入占位符
     * 
     * @param content
     * @return
     */
    private static List<Section> parseInput(String content) {
        List<Section> result = new ArrayList<>();
        
        try{
            for (String json : extractJson(content)) {
                JSONObject jsonObj = JSON.parseObject(json);
                if (null == jsonObj.get("var_name")) {
                    continue;
                }

                Section inputSection = new Section("input", (String) jsonObj.get("name"));
                inputSection.setInputAttr(jsonObj);
                result.add(inputSection);
            }
        } catch (JSONException e){
            // ignore exceptions
        }

        return result;
    }

    /**
     * 提取json信息
     * 
     * @param content
     * @return
     */
    protected static List<String> extractJson(String mixedContent) {
        List<String> result = new ArrayList<>();
        Stack<Integer> stack = new Stack<>();
        Boolean slash = false;
        for (int i = 0; i < mixedContent.length(); i++) {
            if (slash) {
                // 忽略所有的转义字符
                slash = false;
                continue;
            }
            if (mixedContent.charAt(i) == '{') {
                stack.push(i);
            } else if (mixedContent.charAt(i) == '}') {
                if (stack.isEmpty()) {
                    // System.out.println("Invalid JSON string: unmatched closing bracket at
                    // position " + i);
                    continue;
                }
                int start = stack.pop();
                String json = mixedContent.substring(start, i + 1);
                if (stack.size() == 0) {
                    result.add(json);
                }
            } else if (mixedContent.charAt(i) == '\\') {
                slash = true;
            }
        }
        // if (!stack.isEmpty()) {
        // System.out.println("Invalid JSON string: unmatched opening bracket at
        // position " + stack.pop());
        // }
        return result;
    }

}
