package com.laotie.app;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.List;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONObject;

public class WordWriter extends WordParser {
    private JSONObject formValues;

    public static void main(String[] args) {
        String formData = "{\"ent_name\":\"fongwell\",\"ent_code\":\"9160000xxx\",\"farenxingming\":\"laotie\",\"if_crime\":\"是\",\"3_1_dengjizhutigaishu\":\"<div><p>这是第一段</p><p>这是第二段</p><img src=\\\"image/img.jpg\\\"/><p>这是第三段</p></div>\"}";
        try {
            WordWriter writer = new WordWriter("docs/template.docx", JSON.parseObject(formData));
            writer.writeTemplate("docs/output.docx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public WordWriter(String templatePath, JSONObject inputContent) throws IOException {
        super(templatePath);
        this.formValues = inputContent;
    }

    /**
     * 写段落内容
     * 
     * @param paragraph
     * @param content
     * @return
     */
    private int _writeParagraph(XWPFParagraph paragraph) {
        String content = paragraph.getText();
        List<String> jsons = extractJson(content);
        int offset = 0;
        Boolean changed = false;

        if (jsons.size() >= 1) {
            // inline replace
            for (String json : jsons) {
                JSONObject inputObj = JSON.parseObject(json);
                if (null == inputObj.get("var_name") || null == inputObj.get("input_type")) {
                    continue;
                }
                String inputType = inputObj.getString("input_type");
                switch (inputType) {
                    case "text":
                    case "radio":
                        String textInput = formValues.getString(inputObj.getString("var_name"));
                        if (null != textInput) {
                            content = content.replace(json, textInput);
                            changed = true;
                        }
                        break;
                    default:
                        break;
                }
            }

            if (changed) {
                int len = paragraph.getRuns().size();
                // 清除原段落中的内容块，只保留第0个
                for (int i = len - 1; i >= 1; i--) {
                    try {
                        paragraph.removeRun(i);
                    } catch (Exception e) {
                    }
                }
                // 设置新内容
                XWPFRun run = paragraph.getRuns().get(0);
                run.setText(content, 0);
                return offset;
            }
        }
        if (jsons.size() == 1) {
            // whole replace
            int paraIndex = document.getBodyElements().indexOf(paragraph);

            JSONObject inputObj = JSON.parseObject(jsons.get(0));
            if (null == inputObj.get("var_name") || null == inputObj.get("input_type")) {
                return offset;
            }
            String inputType = inputObj.getString("input_type");

            switch (inputType) {
                case "WYSIWYG":
                    String textInput = formValues.getString(inputObj.getString("var_name"));
                    if (null != textInput) {
                        int _offset = _writeWYSIWYParagraphs(textInput, paragraph);
                        if (_offset > 0) {
                            changed = true;
                            document.removeBodyElement(paraIndex + _offset);
                            offset--;
                        }
                        offset += _offset;
                    }
                    break;
                case "table":
                    break;
            }
        }

        return offset;
    }

    /**
     * 写富文本结果，需要转为多行图文混合形式
     * @param HtmlContent
     * @param currPara
     * @return
     */
    private int _writeWYSIWYParagraphs(String HtmlContent, XWPFParagraph currPara) {
        // 使用 Jsoup 解析 HTML
        Document dom = Jsoup.parse(HtmlContent);
        // 遍历所有子元素
        int offset = 0;
        List<Element> tags = dom.select("div").first().children();
        // 因为每次都插入这 cursor 前面的位置，因此倒序输出的结果才是顺序
        Collections.reverse(tags);
        for (Element child : tags) {
            // 处理 <p> 元素
            XmlCursor cursor = currPara.getCTP().newCursor();
            if (child.tagName().equals("p")) {
                String text = child.text();
                if (text.trim().length() == 0) {
                    continue;
                }
                XWPFParagraph newPara = document.insertNewParagraph(cursor);
                newPara.createRun().setText(text.trim());
                currPara = newPara;
                offset++;
            }
            else if (child.tagName().equals("img")) {
            // 处理 <img> 元素
            }
        }
        return offset;
    }

    public void writeTemplate(String saveFile) throws IOException {
        for (int n = 0; n < document.getBodyElements().size(); n++) {
            IBodyElement element = document.getBodyElements().get(n);
            // for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                try {
                    n += _writeParagraph(paragraph);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {

                        String content = cell.getText();
                        for (String json : extractJson(content)) {
                            JSONObject inputObj = JSON.parseObject(json);
                            if (null == inputObj.get("var_name")) {
                                continue;
                            }
                            // TODO: content replace
                            // Object inputInfo = jsonObj.get(inputObj.get("var_name"));
                            // content = content.replace(json, inputInfo.toString());
                        }
                        cell.setText(content);
                    }
                }
            }
        }
        FileOutputStream out;
        try {
            out = new FileOutputStream(saveFile);
            document.write(out);
            out.close();
            document.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
