package com.laotie.app;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Base64;
import java.util.Collections;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.jsoup.safety.Safelist;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;

public class WordWriter extends WordParser {
    private JSONObject formValues;

    public static void main(String[] args) {
        String formData = "{\"ent_name\":\"fongwell\",\"ent_code\":\"9160000xxx\",\"farenxingming\":\"laotie\",\"if_crime\":\"是\",\"3_1_dengjizhutigaishu\":\"<p>This is a paragraph with first line indentation.This is a paragraph with first line indentation.This is a paragraph with first line indentation.This is a paragraph with first line indentation.</p><p><img src=\\\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAbdJREFUSEvtlTFoU1EUhr//JdXBTh3ddHWqICIIoi0U8qI4KLiJi3RSJ4cm6lOTCk7q1E3cBAuKJpFCqyAOiqCDuOrmZqc6SPPeL4l5pTTGl7SNIHjXe8/57vn/c88VQ14acn7WAEk5dAsmJ4dUff5mu8D/FsBRuNerzGmkeUbRwnJbkVSKfiRy6/xMcVyztfcbJXR0Yreb8WtgD2YpqNYnBwbE5cIdoWkFCnWjtpRCHE2NuZl/BewDfqzf77sClwvXjKJO0hXJk7rZeOvo9Kib31+ADgCxCE6p8uxJCu8L4FLhgqW7naAECIBl2VNGtxFHAcucU7X+YL18mQCXwrMW99t+mZfCly0tAGNACkP2RVUb9zZ680eAy8dPmmQeyIHfKb/rmKJHK75SOGhrERj9ZaQjVRrXf/d2egJ8tTjhxHVgJ/BJ8Y4juvX425qxnX3juVylcanXw+wJSErhImIC+KJ87rCip1+7WnOmuJ/Z2gdBewoMVkGr9VbzDzXCtKL6582OjkyTN5t4oDbdCuR/BZnqdUuEzuP4Y2ZkxoH00+oCbDVxGh9U6u3cfw+wXTfvOeyGBfgJ4tMNKLPgx7sAAAAASUVORK5CYII=\\\" alt=\\\"\\\" width=\\\"141\\\" height=\\\"141\\\" /><img src=\\\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAbdJREFUSEvtlTFoU1EUhr//JdXBTh3ddHWqICIIoi0U8qI4KLiJi3RSJ4cm6lOTCk7q1E3cBAuKJpFCqyAOiqCDuOrmZqc6SPPeL4l5pTTGl7SNIHjXe8/57vn/c88VQ14acn7WAEk5dAsmJ4dUff5mu8D/FsBRuNerzGmkeUbRwnJbkVSKfiRy6/xMcVyztfcbJXR0Yreb8WtgD2YpqNYnBwbE5cIdoWkFCnWjtpRCHE2NuZl/BewDfqzf77sClwvXjKJO0hXJk7rZeOvo9Kib31+ADgCxCE6p8uxJCu8L4FLhgqW7naAECIBl2VNGtxFHAcucU7X+YL18mQCXwrMW99t+mZfCly0tAGNACkP2RVUb9zZ680eAy8dPmmQeyIHfKb/rmKJHK75SOGhrERj9ZaQjVRrXf/d2egJ8tTjhxHVgJ/BJ8Y4juvX425qxnX3juVylcanXw+wJSErhImIC+KJ87rCip1+7WnOmuJ/Z2gdBewoMVkGr9VbzDzXCtKL6582OjkyTN5t4oDbdCuR/BZnqdUuEzuP4Y2ZkxoH00+oCbDVxGh9U6u3cfw+wXTfvOeyGBfgJ4tMNKLPgx7sAAAAASUVORK5CYII=\\\" alt=\\\"\\\" width=\\\"141\\\" height=\\\"141\\\" /><img src=\\\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAbdJREFUSEvtlTFoU1EUhr//JdXBTh3ddHWqICIIoi0U8qI4KLiJi3RSJ4cm6lOTCk7q1E3cBAuKJpFCqyAOiqCDuOrmZqc6SPPeL4l5pTTGl7SNIHjXe8/57vn/c88VQ14acn7WAEk5dAsmJ4dUff5mu8D/FsBRuNerzGmkeUbRwnJbkVSKfiRy6/xMcVyztfcbJXR0Yreb8WtgD2YpqNYnBwbE5cIdoWkFCnWjtpRCHE2NuZl/BewDfqzf77sClwvXjKJO0hXJk7rZeOvo9Kib31+ADgCxCE6p8uxJCu8L4FLhgqW7naAECIBl2VNGtxFHAcucU7X+YL18mQCXwrMW99t+mZfCly0tAGNACkP2RVUb9zZ680eAy8dPmmQeyIHfKb/rmKJHK75SOGhrERj9ZaQjVRrXf/d2egJ8tTjhxHVgJ/BJ8Y4juvX425qxnX3juVylcanXw+wJSErhImIC+KJ87rCip1+7WnOmuJ/Z2gdBewoMVkGr9VbzDzXCtKL6582OjkyTN5t4oDbdCuR/BZnqdUuEzuP4Y2ZkxoH00+oCbDVxGh9U6u3cfw+wXTfvOeyGBfgJ4tMNKLPgx7sAAAAASUVORK5CYII=\\\" alt=\\\"\\\" width=\\\"141\\\" height=\\\"141\\\" /></p>\\n"
                + //
                "<p>test</p>\\n" + //
                "<p>test2</p>\\n" + //
                "<p>test31</p>\",\"z_table\":{\"columns\":[[\"aaaaa\",\"111111\",\"10\"],[\"bbbbb\",\"222222\",\"20\"],[\"ccccc\",\"333333\",\"30\"]]}}";
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
            String replacedContent = _writeSimpleText(content);
            if (!replacedContent.equalsIgnoreCase(content)) {
                changed = true;
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
                run.setText(replacedContent, 0);
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
                    JSONObject tableData = formValues.getJSONObject(inputObj.getString("var_name"));
                    if (null != tableData) {
                        int _offset = _writeNewTable(inputObj, tableData, paragraph);
                        if (_offset > 0) {
                            changed = true;
                            document.removeBodyElement(paraIndex + _offset);
                            offset--;
                        }
                        offset += _offset;
                    }
                    break;
            }
        }

        return offset;
    }

    /**
     * 替换简单文本
     * 
     * @param content
     * @return
     */
    private String _writeSimpleText(String content) {
        // Boolean changed = false;
        List<String> jsons = extractJson(content);
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
                        // changed = true;
                    }
                    break;
                default:
                    break;
            }
        }
        return content;
    }

    /**
     * 写富文本结果，需要转为多行图文混合形式
     * 
     * @param HtmlContent
     * @param currPara
     * @return
     */
    private int _writeWYSIWYParagraphs(String HtmlContent, XWPFParagraph currPara) {
        // 使用 Jsoup 解析 HTML，只保留<p><img>标签
        HtmlContent = Jsoup.clean(HtmlContent,
                Safelist.none().addTags("p", "img").addAttributes("img", "alt", "height", "src", "width"));
        Document dom = Jsoup.parse(HtmlContent);
        // 遍历所有子元素
        int offset = 0;
        List<Element> tags = dom.select("p");
        // 因为获得的 cursor 在段落前面的位置，并且没有找到获得段落后面的方法，因此倒序插入就是顺序
        Collections.reverse(tags);
        for (Element ptag : tags) {
            XmlCursor cursor = currPara.getCTP().newCursor();
            XWPFParagraph newPara = document.insertNewParagraph(cursor);
            for (Node childNode : ptag.childNodes()) {
                if (childNode instanceof TextNode) {
                    // 段落中的文本
                    TextNode textNode = (TextNode) childNode;
                    String text = textNode.text();
                    if (text.trim().length() == 0) {
                        continue;
                    }
                    XWPFRun r = newPara.createRun();
                    // 设置文本字体和首行缩进
                    r.setText(text.trim());
                    r.setFontFamily("宋体");
                    r.setFontSize(12);
                    newPara.setFirstLineIndent(600);
                } else if (childNode instanceof Element) {
                    // 段落中的图片
                    Element pcontent = (Element) childNode;
                    if (!pcontent.tagName().equals("img")) {
                        continue;
                    }
                    int width = Integer.valueOf(pcontent.attr("width"));
                    int height = Integer.valueOf(pcontent.attr("height"));
                    if (width == 0 || height == 0) {
                        continue;
                    }
                    String src = pcontent.attr("src");
                    String[] photoData = src.split(";base64,", 2);
                    if (photoData.length <= 1) {
                        continue;
                    }
                    String base64Image = photoData[1];
                    String photoType = photoData[0].replace("data:image/", "");

                    XWPFRun r = newPara.createRun();
                    byte[] imageBytes = Base64.getDecoder().decode(base64Image);
                    ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);
                    try {
                        r.addPicture(bis, _getPictureType(photoType), "image." + photoType,
                                Units.toEMU(width), Units.toEMU(height));
                        newPara.setAlignment(ParagraphAlignment.CENTER);
                    } catch (InvalidFormatException | IOException e) {
                        e.printStackTrace();
                    }
                    r.addBreak();
                }
            }
            currPara = newPara;
            offset++;
        }
        return offset;
    }

    /**
     * 获得word中的图片类型编码
     * @param typename
     * @return
     */
    private int _getPictureType(String typename) {
        switch (typename) {
            case "png":
                return XWPFDocument.PICTURE_TYPE_PNG;
            case "jpeg":
            case "jpg":
                return XWPFDocument.PICTURE_TYPE_JPEG;
            case "gif":
                return XWPFDocument.PICTURE_TYPE_GIF;
            case "bmp":
                return XWPFDocument.PICTURE_TYPE_BMP;
            default:
                return XWPFDocument.PICTURE_TYPE_PNG;
        }
    }

    /**
     * 插入一个新的表格
     * @param inputObj
     * @param tableData
     * @param currPara
     * @return
     */
    private int _writeNewTable(JSONObject inputObj, JSONObject tableData, XWPFParagraph currPara) {
        XmlCursor cursor = currPara.getCTP().newCursor();
        XWPFTable newTable = document.insertNewTbl(cursor);

        // 使用 FastJSON 解析 JSON 数组
        JSONArray dataByColumns = tableData.getJSONArray("columns");
        if (null == dataByColumns || dataByColumns.size() == 0) {
            return 0;
        }

        JSONArray header = inputObj.getJSONObject("input_des").getJSONArray("columns");
        XWPFTableRow currRow = newTable.getRow(0);
        currRow.getCell(0).setText(header.getJSONObject(0).getString("name"));
        for (int i = 1; i < header.size(); i++) {
            currRow.createCell().setText(header.getJSONObject(i).getString("name"));
        }

        // 遍历 JSON 数组
        for (int i = 0; i < dataByColumns.size(); i++) {
            JSONArray innerArray = dataByColumns.getJSONArray(i);
            currRow = newTable.getRow(i + 1);
            if (null == currRow) {
                currRow = newTable.createRow();
            }
            for (int j = 0; j < innerArray.size(); j++) {
                currRow.getCell(j).setText(innerArray.getString(j));
            }
        }

        JSONArray footer = inputObj.getJSONObject("input_des").getJSONArray("rows");
        currRow = newTable.createRow();
        int mergedSize = 0;
        for (int i = 0; i < footer.size(); i++) {
            JSONObject footerCell = footer.getJSONObject(i);
            switch (footerCell.getString("type")) {
                case "const":
                    currRow.getCell(i + mergedSize).setText(footerCell.getString("content"));
                    int colspan = footerCell.getIntValue("colspan");
                    if (colspan > 0) {
                        currRow.getCell(i).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                        currRow.getCell(i + colspan - 1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                        mergedSize += colspan - 1;
                    }
                    break;
                case "sum":
                    double sum = 0;
                    for (int j = 0; j < dataByColumns.size(); j++) {
                        sum += dataByColumns.getJSONArray(j).getDouble(i + mergedSize);
                    }
                    if (sum == (long) sum) {
                        currRow.getCell(i + mergedSize).setText(String.format("%d", (long) sum));
                    } else {
                        currRow.getCell(i + mergedSize).setText(String.format("%s", sum));
                    }
                default:
                    break;
            }
        }
        return 1;
    }

    /**
     * 向模板中填充数据
     * @param saveFile
     * @throws IOException
     */
    public void writeTemplate(String saveFile) throws IOException {
        for (int n = 0; n < document.getBodyElements().size(); n++) {
            IBodyElement element = document.getBodyElements().get(n);
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                n += _writeParagraph(paragraph);
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        String content = cell.getText();
                        content = _writeSimpleText(content);
                        cell.setText(content);
                    }
                }
            }
        }
        FileOutputStream out;
        out = new FileOutputStream(saveFile);
        document.write(out);
        out.close();
        document.close();
    }

}
