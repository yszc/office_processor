package com.laotie.app;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;

public class WordWriter extends WordParser {
    private JSONObject formValues;

    public static void main(String[] args) {
        String formData = "{\n" + //
                "  \"0_ent_name\": \"广州芳禾数据有限公司\",\n" + //
                "}";
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
     * @return
     */
    private int _writeParagraph(XWPFParagraph paragraph) {
        String content = paragraph.getText();
        List<String> jsons = extractJson(content);
        int offset = 0;
        Boolean changed = false;
        if (jsons.isEmpty()) {
            return offset;
        }
        Set<String> inputTypes = _getInputTypes(jsons);
        if (inputTypes.contains("table") || inputTypes.contains("WYSIWYG")) {
            // whole replace
            int paraIndex = document.getBodyElements().indexOf(paragraph);
            int _offset = 0;

            JSONObject inputObj = JSON.parseObject(jsons.get(0));
            if (null == inputObj.get("var_name") || null == inputObj.get("input_type")) {
                return offset;
            }
            String inputType = inputObj.getString("input_type");

            switch (inputType) {
                case "WYSIWYG":
                    String textInput = formValues.getString(inputObj.getString("var_name"));
                    if (null != textInput) {
                        _offset = _writeWYSIWYParagraphs(textInput, paragraph);
                        offset += _offset;
                    }
                    document.removeBodyElement(paraIndex + _offset);
                    offset--;
                    break;
                case "table":
                    JSONArray tableData = formValues.getJSONArray(inputObj.getString("var_name"));
                    _offset = _writeNewTable(inputObj, tableData, paragraph);
                    offset += _offset;
                    document.removeBodyElement(paraIndex + _offset);
                    offset--;
                    break;
            }
        } else {
            // inline replace
            List<String> replacedContent = _writeSimpleTextAndMarkUnderline(content, true);
            // 清除原段落中的内容块，只保留第0个以保持样式
            while (paragraph.getRuns().size() > 1) {
                paragraph.removeRun(1);
            }
            // 获取字体
            String fontFamily = paragraph.getRuns().get(0).getFontFamily();
            Double fontSize = paragraph.getRuns().get(0).getFontSizeAsDouble();
            for (int i = 0; i < replacedContent.size(); i++) {
                // 设置新内容
                XWPFRun run = null;
                if (i < paragraph.getRuns().size()) {
                    run = paragraph.getRuns().get(i);
                } else {
                    run = paragraph.createRun();
                    run.setFontFamily(fontFamily);
                    run.setFontSize(fontSize);
                }
                String runContent = replacedContent.get(i);
                if (runContent.length() > 4 && runContent.substring(0, 2).equals("$$")
                        && runContent.substring(runContent.length() - 2).equals("$$")) {
                    runContent = runContent.substring(2, runContent.length() - 2);
                    run.setUnderline(UnderlinePatterns.SINGLE);
                }
                run.setText(runContent, 0);
            }
            return offset;
        }

        return offset;
    }

    /**
     * 获得所有占位符类型
     * 
     * @param jsons
     * @return
     */
    private Set<String> _getInputTypes(List<String> jsons) {
        Set<String> inputTypes = new HashSet<>();
        for (String json : jsons) {
            JSONObject inputObj = JSON.parseObject(json);
            if (null != inputObj.get("input_type")) {
                inputTypes.add(inputObj.getString("input_type"));
            }
        }
        return inputTypes;
    }

    /**
     * 替换简单文本，并标记下划线
     * 
     * @param content
     * @param pure
     * @return
     */
    private List<String> _writeSimpleTextAndMarkUnderline(String content, Boolean pure) {
        // Boolean changed = false;
        List<String> jsons = extractJson(content);
        List<String> contentRuns = new ArrayList<>();
        for (String json : jsons) {
            String[] splited = content.split(Pattern.quote(json), 2);
            contentRuns.add(splited[0]);
            JSONObject inputObj = JSON.parseObject(json);
            String replacement = "";
            String inputType = inputObj.getString("input_type");
            String varName = inputObj.getString("var_name");
            if (null != varName && null != inputType) {
                switch (inputType) {
                    case "text":
                    case "radio":
                    case "date":
                        replacement = formValues.getString(varName);
                        if (!replacement.isEmpty()) {
                            replacement = "$$" + replacement + "$$";
                        }
                        break;
                    default:
                        break;
                }
            }
            if (!pure && replacement.isEmpty()) {
                replacement = json;
            }
            if (replacement.isEmpty()) {
                replacement = "$$         $$";
            }
            contentRuns.add(replacement);
            content = splited[1];
        }
        if (!content.trim().isEmpty()) {
            contentRuns.add(content);
        }
        return contentRuns;
    }

    /**
     * 替换简单文本
     *
     * @param content
     * @return
     */
    private String _writeSimpleText(String content) {
        return _writeSimpleText(content, true);
    }

    /**
     * 替换简单文本
     *
     * @param content
     * @param pure    是否过滤掉未填充的占位符
     * @return
     */
    private String _writeSimpleText(String content, Boolean pure) {
        // Boolean changed = false;
        List<String> jsons = extractJson(content);
        for (String json : jsons) {
            JSONObject inputObj = JSON.parseObject(json);
            if (null == inputObj.get("var_name") || null == inputObj.get("input_type")) {
                if (pure) {
                    content = content.replace(json, "");
                }
                continue;
            }
            String inputType = inputObj.getString("input_type");
            switch (inputType) {
                case "text":
                case "radio":
                case "date":
                    String textInput = formValues.getString(inputObj.getString("var_name"));
                    if (null != textInput) {
                        content = content.replace(json, textInput);
                        // changed = true;
                    }
                    break;
                default:
                    break;
            }
            if (pure) {
                content = content.replace(json, "");
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

                    try {
                        int width = Integer.valueOf("0" + pcontent.attr("width"));
                        int height = Integer.valueOf("0" + pcontent.attr("height"));
                        width = width == 0 ? 300 : width;
                        height = height == 0 ? 300 : height;

                        int[] adapted = _getImageZoom(width, height, 400, 600);
                        width = adapted[0];
                        height = adapted[1];

                        XWPFRun r = newPara.createRun();
                        String src = pcontent.attr("src");
                        String photoType = _getImageType(src);
                        InputStream bis = _getImageByteStream(src);
                        r.addPicture(bis, _getPictureType(photoType), "image." + photoType,
                                Units.toEMU(width), Units.toEMU(height));
                        newPara.setAlignment(ParagraphAlignment.CENTER);
                        r.addBreak();
                    } catch (InvalidFormatException | IOException e) {
                        e.printStackTrace();
                    }
                }
            }
            currPara = newPara;
            offset++;
        }
        return offset;
    }

    /**
     * 获取图片缩放适配的宽高
     * 
     * @param originalWidth  原宽度
     * @param originalHeight 原高度
     * @param maxWidth       最大宽度
     * @param maxHeight      最大高度
     * @return 宽高
     */
    private static int[] _getImageZoom(int originalWidth, int originalHeight, int maxWidth, int maxHeight) {
        int adaptedWidth = originalWidth;
        int adaptedHeight = originalHeight;

        // 如果图片的宽大于最大宽度或高大于最大高度，则需要适配
        if (originalWidth > maxWidth || originalHeight > maxHeight) {
            // 计算宽高比例
            double aspectRatio = (double) originalWidth / originalHeight;

            // 根据最大宽度适配，然后计算相应的高度
            adaptedWidth = maxWidth;
            adaptedHeight = (int) (adaptedWidth / aspectRatio);

            // 如果适配后的高度超过最大高度，则使用最大高度适配，然后计算相应的宽度
            if (adaptedHeight > maxHeight) {
                adaptedHeight = maxHeight;
                adaptedWidth = (int) (adaptedHeight * aspectRatio);
            }
        }

        return new int[] { adaptedWidth, adaptedHeight };
    }

    /**
     * 获得图片的类型
     * 
     * @param src
     * @return
     * @throws IOException
     */
    private String _getImageType(String src) throws IOException {
        String defaultType = "png";
        if (src.indexOf("http", 0) == 0) {
            URL url = new URL(src);
            URLConnection connection = url.openConnection();
            String contentType = connection.getContentType();
            if (contentType.indexOf("image/") == 0) {
                return contentType.replace("image/", "");
            }
            return "png";
        } else {
            String[] photoData = src.split(";base64,", 2);
            if (photoData.length > 1 && photoData[0].indexOf("data:image/") == 0) {
                return photoData[0].replace("data:image/", "");
            }
        }
        return defaultType;
    }

    /**
     * 根据img标签的src属性，获取图片字节流
     * 
     * @param src
     * @return
     * @throws MalformedURLException
     * @throws IOException
     */
    private InputStream _getImageByteStream(String src) throws MalformedURLException, IOException {
        if (src.indexOf("http", 0) == 0) {
            return _getRemoteImageByteStream(src);
        } else {
            return _getBase64ImageByteStream(src);
        }
    }

    /**
     * 根据base64获取图片字节流
     * 
     * @param src
     * @return
     */
    private InputStream _getBase64ImageByteStream(String src) {
        String[] photoData = src.split(";base64,", 2);
        if (photoData.length <= 1) {
            return null;
        }
        String base64Image = photoData[1];

        byte[] imageBytes = Base64.getDecoder().decode(base64Image);
        ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);
        return bis;
    }

    /**
     * 根据URL获取图片字节流
     * 
     * @param src
     * @return
     * @throws MalformedURLException
     * @throws IOException
     */
    private InputStream _getRemoteImageByteStream(String src) throws MalformedURLException, IOException {
        return new URL(src).openStream();
    }

    /**
     * 获得word中的图片类型编码
     * 
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
     * 
     * @param inputObj
     * @param tableData
     * @param currPara
     * @return
     */
    private int _writeNewTable(JSONObject inputObj, JSONArray tableData, XWPFParagraph currPara) {
        XmlCursor cursor = currPara.getCTP().newCursor();
        XWPFTable newTable = document.insertNewTbl(cursor);

        // 设置表格的宽度为页面的全宽
        CTTblPr tblPr = newTable.getCTTbl().getTblPr();
        if (tblPr == null) {
            tblPr = newTable.getCTTbl().addNewTblPr();
        }
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        tblWidth.setW(5000);
        tblWidth.setType(STTblWidth.PCT);

        // 写表头
        JSONArray header = inputObj.getJSONObject("input_des").getJSONArray("columns");
        XWPFTableRow currRow = newTable.getRow(0);
        for (int i = 0; i < header.size(); i++) {
            XWPFTableCell cell = null != currRow.getCell(i) ? currRow.getCell(i) : currRow.createCell();
            cell.setText(header.getJSONObject(i).getString("name"));
            XWPFRun r = cell.getParagraphs().get(0).getRuns().get(0);
            // r.setBold(true);
            r.setFontFamily("仿宋_GB2312");
            r.setFontSize(13);
        }

        if (null == tableData || tableData.size() == 0) {
            return 1;
        }
        // 遍历 JSON 数组，写表体
        for (int i = 0; i < tableData.size(); i++) {
            JSONArray innerArray = tableData.getJSONArray(i);
            currRow = newTable.getRow(i + 1);
            if (null == currRow) {
                currRow = newTable.createRow();
            }
            for (int j = 0; j < innerArray.size(); j++) {
                currRow.getCell(j).setText(innerArray.getString(j));
                XWPFRun r = currRow.getCell(j).getParagraphs().get(0).getRuns().get(0);
                r.setFontFamily("仿宋_GB2312");
                r.setFontSize(13);
            }
        }

        // 写表脚
        JSONArray footer = inputObj.getJSONObject("input_des").getJSONArray("rows");
        currRow = newTable.createRow();
        int mergedSize = 0;
        for (int i = 0; i < footer.size(); i++) {
            JSONObject footerCell = footer.getJSONObject(i);
            int currCellIndex = i + mergedSize;
            XWPFTableCell currCell = currRow.getCell(currCellIndex);
            switch (footerCell.getString("type")) {
                case "const":
                    currCell.setText(footerCell.getString("content"));
                    int colspan = footerCell.getIntValue("colspan");
                    if (colspan > 1) {
                        currRow.getCell(currCellIndex).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                        currRow.getCell(currCellIndex + colspan - 1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                        mergedSize += colspan - 1;
                    }
                    break;
                case "sum":
                    double sum = 0;
                    for (int j = 0; j < tableData.size(); j++) {
                        sum += tableData.getJSONArray(j).getDouble(i + mergedSize);
                    }
                    if (sum == (long) sum) {
                        currCell.setText(String.format("%d", (long) sum));
                    } else {
                        currCell.setText(String.format("%s", sum));
                    }
                    break;
                default:
                    currCell.setText(" ");
                    break;
            }

            XWPFRun r = currCell.getParagraphs().get(0).getRuns().get(0);
            r.setFontFamily("仿宋_GB2312");
            r.setFontSize(13);
        }
        return 1;
    }

    /**
     * 设置table种的单元格内容，并保留原样式
     * 
     * @param cell
     * @param text
     */
    private void _setCellText(XWPFTableCell cell, String text) {
        // 保留第一个paragraph第一个run以保持模板样式
        while (cell.getParagraphs().size() > 1) {
            cell.removeParagraph(1);
        }
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        XWPFParagraph par = paragraphs.isEmpty() ? cell.addParagraph() : paragraphs.get(0);
        while (par.getRuns().size() > 1) {
            par.removeRun(1);
        }
        XWPFRun run = par.getRuns().size() == 0 ? par.createRun() : par.getRuns().get(0);
        run.setText(text, 0);
    }

    /**
     * 填充document对象中的占位信息
     */
    public void writeDocument() {
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
                        _setCellText(cell, content);
                    }
                }
            }
        }
    }

    /**
     * 向模板中填充数据，并写文件
     * 
     * @param saveFile
     * @throws IOException
     */
    public void writeTemplate(String saveFile) throws IOException {
        writeDocument();
        FileOutputStream out;
        out = new FileOutputStream(saveFile);
        document.write(out);
        out.close();
        document.close();
    }

}
