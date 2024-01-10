package com.laotie.app;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.*;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
        // 编译正则表达式
        Pattern pattern = Pattern.compile("\\{([^}]+)\\}");
        // 查找占位符
        Matcher matcher = pattern.matcher(content);
        List<Section> result = new ArrayList<>();
        while (matcher.find()) {
            String match = matcher.group(0); // 获取第一个捕获组的内容
            Section paraSection = new Section("input", match);
            result.add(paraSection);
        }
        return result;
    }

    private static class Section {
        private String type;
        private String prefix = "";
        private String name;
        private List<Section> children;

        /**
         * 创建一个子节点
         * 
         * @param type
         * @param name
         * @return
         */
        public Section(String type, String name) {
            this.type = type;
            this.name = name;
            this.children = new ArrayList<>();
        }

        public String toJson() throws JsonProcessingException {
            ObjectMapper mapper = new ObjectMapper();
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(this);
        }

        public static Section fromJson(String jsonResult) throws JsonMappingException, JsonProcessingException {
            ObjectMapper mapper = new ObjectMapper();
            return mapper.readValue(jsonResult, Section.class);
        }

        /**
         * 获取非空的结构
         * 
         * @return
         */
        public Section toNoneEmpty() {
            return filterSection(this);
        }

        /**
         * 过滤所有空标题
         * 
         * @param root
         * @return
         */
        private static Section filterSection(Section root) {
            if (null == root.getChildren()) {
                return root;
            }
            for (Section child : root.getChildren()) {
                filterSection(child);
            }
            root.setChildren(new ArrayList<>(CollectionUtils.select(root.getChildren(),
                    child -> null != child && (child.getType() == "input" || child.getChildren().size() > 0))));
            return root;
        }

        public String getType() {
            return type;
        }

        public void setType(String type) {
            this.type = type;
        }

        public String getPrefix() {
            return prefix;
        }

        public void setPrefix(String prefix) {
            this.prefix = prefix;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public List<Section> getChildren() {
            return children;
        }

        public void setChildren(List<Section> children) {
            this.children = children;
        }

    }
}
