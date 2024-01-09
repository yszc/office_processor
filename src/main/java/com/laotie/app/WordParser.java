package com.laotie.app;


import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.*;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordParser {

    public static void main(String[] args) {
        String filePath = "/workspaces/wordparser/template.docx";
        try (FileInputStream fis = new FileInputStream(filePath)) {
            XWPFDocument document = new XWPFDocument(fis);

            Section root = parseDocument(document);
            root = filterSection(root);

            ObjectMapper mapper = new ObjectMapper();
            String jsonResult = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(root.getChildren());
            System.out.println(jsonResult);

            List<Section> sections = mapper.readValue(jsonResult, new TypeReference<List<Section>>() {});
            Section rootBack = new Section();
            rootBack.setChildren(sections);
            System.out.println(rootBack);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Section filterSection(Section root){
        if (null == root.getChildren()){
            return root;
        }
        for (Section child: root.getChildren()){
            filterSection(child);
        }
        root.setChildren(new ArrayList<>(CollectionUtils.select(root.getChildren(), child -> null != child && (child.getType()=="input" || child.getChildren().size() > 0))));
        return root;
    }

    private static int getHeadingIndent(String heading){
        if (heading.startsWith("heading")){
            return Integer.parseInt(heading.replace("heading", "").trim());
        }
        return 0;
    }

    private static Section createSection(String type, String name, List<Section> children){
        Section paraSection = new Section();
        paraSection.setType(type);
        paraSection.setName(name);
        paraSection.setChildren(children);
        return paraSection;     
    }

    private static Section parseDocument(XWPFDocument document) {
        Stack<Section> stack = new Stack<>();
        Section root = createSection("root", "root", new ArrayList<>());
        stack.push(root);

        XWPFStyles style_sheet = document.getStyles();
        // 编译正则表达式
        Pattern pattern = Pattern.compile("\\{([^}]+)\\}");

        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                String paraText = paragraph.getText();
                String styleID = paragraph.getStyleID();
                if (styleID != null ) {
                    String styleName = style_sheet.getStyle(styleID).getName().toLowerCase();
                    if(styleName.startsWith("heading")){
                        int indent = getHeadingIndent(styleName);
                        if (indent==0){
                            continue;
                        }
                        while (stack.size()>indent){
                            stack.pop();
                        }
                        Section paraSection = createSection("title", paraText, new ArrayList<>());
                        if (!stack.isEmpty()){
                            stack.peek().getChildren().add(paraSection);
                        }
                        String parentPrefix = stack.peek().getPrefix();
                        List<Section> brothers = stack.peek().getChildren();
                        if (parentPrefix.length()>0){
                            paraSection.setPrefix(parentPrefix + "." + brothers.size());
                        }else{
                            paraSection.setPrefix(String.valueOf(brothers.size()));
                        }
                        stack.push(paraSection);
                    }
                } else {
                    // 创建 Matcher 对象
                    Matcher matcher = pattern.matcher(paraText);
                    // 查找匹配
                    while (matcher.find()) {
                        String match = matcher.group(0); // 获取第一个捕获组的内容
                        Section paraSection = createSection("input", match, new ArrayList<>());
                        stack.peek().getChildren().add(paraSection);
                    }
                }
            }
        }
        return root;
    }


    private static class Section {
        private String type;
        private String prefix="";
        private String name;
        private List<Section> children;

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
