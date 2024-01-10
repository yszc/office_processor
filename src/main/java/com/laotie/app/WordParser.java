package com.laotie.app;

import org.apache.poi.xwpf.usermodel.*;

import com.fasterxml.jackson.core.JsonProcessingException;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;

public class WordParser {
    private XWPFDocument document;
    private XWPFStyles style_sheet;
    private Section nestedRoot;

    public static void main(String[] args) {
        String filePath = "/workspaces/wordparser/docs/template.docx";
        try {
            WordParser wordParser = new WordParser(filePath);
            Section root = wordParser.parseTemplate();

            String jsonResult = root.toNoneEmpty().toJson();
            System.out.println(jsonResult);

            Section rootBack = Section.fromJson(jsonResult);
            System.out.println(rootBack);
        } catch (Exception e) {

        }
    }

    public WordParser(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath)) {
            this.document = new XWPFDocument(fis);
            // 样式表
            this.style_sheet = document.getStyles();
            this.nestedRoot = new Section("root", "root");
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 解析模板，获得标题和对应的输入框信息
     * 
     * @param document
     * @return
     */
    public Section parseTemplate() {
        Stack<Section> stack = new Stack<>();
        stack.push(nestedRoot);

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
                paraSection.setPrefix(parentPrefix + "." + brothers.size());
            } else {
                paraSection.setPrefix(String.valueOf(brothers.size()));
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

    // 解析输入占位符
    private static List<Section> parseInput(String content) {
        List<Section> result = new ArrayList<>();

        for (String json : extractJson(content)) {
            try {
                Section inputSection;
                inputSection = Section.fromJson(json);
                inputSection.setType("input");
                result.add(inputSection);
            } catch (JsonProcessingException e) {
                e.printStackTrace();
            }
        }

        return result;
    }

    /**
     * 提取json信息
     * 
     * @param content
     * @return
     */
    private static List<String> extractJson(String mixedContent) {
        List<String> result = new ArrayList<>();
        Stack<Integer> stack = new Stack<>();
        for (int i = 0; i < mixedContent.length(); i++) {
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
            }
        }
        // if (!stack.isEmpty()) {
        // System.out.println("Invalid JSON string: unmatched opening bracket at
        // position " + stack.pop());
        // }
        return result;
    }

}
