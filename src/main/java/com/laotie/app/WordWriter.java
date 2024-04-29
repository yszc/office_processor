package com.laotie.app;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.*;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
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
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
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
        if (jsons.isEmpty()) {
            return offset;
        }
        Set<String> inputTypes = _getInputTypes(jsons);
        if (inputTypes.contains("table") || inputTypes.contains("WYSIWYG")||inputTypes.contains("checkbox") || inputTypes.contains("file_list")) {
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
                case "checkbox":
                    JSONArray checked = formValues.getJSONArray(inputObj.getString("var_name"));
                    _offset = _writeCheckbox(inputObj, checked, paragraph);
                    offset += _offset;
                    document.removeBodyElement(paraIndex + _offset);
                    offset--;
                    break;
                case "file_list":
                    JSONArray fileList = formValues.getJSONArray(inputObj.getString("var_name"));
                    _offset = _writeFileListCard(fileList, paragraph);
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
            String fontFamily = _getFontFamily(paragraph);
            Double fontSize = _getFontSize(paragraph);
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
                    case "select":
                    case "date":
                        replacement = formValues.getString(varName);
                        if (!StringUtils.isEmpty(replacement)) {
                            replacement = "$$" + replacement + "$$";
                        }
                        break;
                    default:
                        break;
                }
            }
            if (!pure && StringUtils.isEmpty(replacement)) {
                replacement = json;
            }
            if (StringUtils.isEmpty(replacement)) {
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
     * @param pure    是否过滤掉未填充的占位符，false 会保留 json 在结果中
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
                case "select":
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
     * 获得段落字体
     */
    private String _getFontFamily(XWPFParagraph currPara){
        try{
            String fontFamily = currPara.getRuns().get(0).getFontFamily();
            if (!Objects.isNull(fontFamily)){
                return fontFamily;
            }
        }catch(Exception e){}
        return "宋体";
    }

    /**
     * 获得段落字体大小
     */
    private Double _getFontSize(XWPFParagraph currPara){
        try{
            Double fontSizeAsDouble = currPara.getRuns().get(0).getFontSizeAsDouble();
            if (!Objects.isNull(fontSizeAsDouble)){
                return fontSizeAsDouble;
            }
        }catch(Exception e){}
        return 13.0;
    }

    /**
     * 设置段落字体
     */
    private void _setParagraphFontFamily(XWPFParagraph currPara, String fontFamily){
        try{
            List<XWPFRun> runs = currPara.getRuns();;
            if (runs.size()==0){
                runs.add(currPara.createRun());
            }
            for(int i=0; i<runs.size(); i++){
                runs.get(i).setFontFamily(fontFamily);
            }
        }catch(Exception e){}
    }

    /**
     * 设置段落字体大小
     */
    private void _setParagraphFontSize(XWPFParagraph currPara, Double fontSize){
        try{
            List<XWPFRun> runs = currPara.getRuns();;
            if (runs.size()==0){
                runs.add(currPara.createRun());
            }
            for(int i=0; i<runs.size(); i++){
                runs.get(i).setFontSize(fontSize.intValue());
            }
        }catch(Exception e){}
    }

    /**
     * 写段落中的文件列表（只显示文件名）
     *
     * @param filelist
     * @param currPara
     * @return
     */
    private int _writeFileListCard(JSONArray filelist, XWPFParagraph currPara) {
        int offset = 0;
        if(null == filelist){
            return offset;
        }
        String toHtml = "";
        for(int i = 0; i < filelist.size(); i++){
            JSONObject fileObj = filelist.getJSONObject(i);
            String fileName = fileObj.getString("name");
            toHtml += "<p>"+ fileName + "</p>";
        }
        offset = _writeWYSIWYParagraphs(toHtml, currPara);
        return offset;
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
        List<Element> tags = dom.select("p,img");
        // 因为获得的 cursor 在段落前面的位置，并且没有找到获得段落后面的方法，因此倒序插入就是顺序
        Collections.reverse(tags);
        String fontFamily = _getFontFamily(currPara);
        Double fontSize = _getFontSize(currPara);
        for (Element tagElm : tags) {
            XmlCursor cursor = currPara.getCTP().newCursor();
            XWPFParagraph newPara = document.insertNewParagraph(cursor);
            _writePTag(tagElm, newPara, fontFamily, fontSize);
            _writeImageTag(tagElm, newPara);
            currPara = newPara;
            offset++;
        }
        return offset;
    }

    /**
     * 跟进 p 标签，插入多行图文混合形式
     *
     * @param tagElm
     * @param newPara
     * @return
     */
    private void _writePTag(Element tagElm, XWPFParagraph newPara, String fontFamily, Double fontSize) {
        if (!tagElm.tagName().equals("p")) {
            return;
        }
        for (Node childNode : tagElm.childNodes()) {
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
                r.setFontFamily(fontFamily);
                r.setFontSize(fontSize);
                newPara.setFirstLineIndent(600);
            } else if (childNode instanceof Element) {
                // 段落中的标签
                Element pcontent = (Element) childNode;
                _writePTag(pcontent, newPara, fontFamily, fontSize);
                _writeImageTag(pcontent, newPara);
            }
        }
    }

    /**
     * 根据 img 标签在段落中插入图片
     *
     * @param tagElm
     * @param newPara
     * @return
     */
    private void _writeImageTag(Element tagElm, XWPFParagraph newPara){
        if (!tagElm.tagName().equals("img")) {
            return;
        }
        try {
            int width = Integer.valueOf("0" + tagElm.attr("width"));
            int height = Integer.valueOf("0" + tagElm.attr("height"));
            width = width == 0 ? 300 : width;
            height = height == 0 ? 300 : height;

            int[] adapted = _getImageZoom(width, height, 400, 600);
            width = adapted[0];
            height = adapted[1];

            XWPFRun r = newPara.createRun();
            String src = tagElm.attr("src");
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
     * 插入多选项
     * @param currPara
     * @return
     */
    private int _writeCheckbox(JSONObject inputObj, JSONArray checked, XWPFParagraph currPara) {
        int offset = 0;
        if (checked == null){
            checked = new JSONArray();
        }
        String fontFamily = _getFontFamily(currPara);
        Double fontSize = _getFontSize(currPara);
        JSONArray allTheOptions = inputObj.getJSONObject("input_des").getJSONArray("options");
        // write each option single line, and if it checked then write it as checked
        for (int i = 0; i < allTheOptions.size(); i++) {
            XmlCursor cursor = currPara.getCTP().newCursor();
            XWPFParagraph newPara = document.insertNewParagraph(cursor);
            offset++;
            newPara.setFirstLineIndent(currPara.getFirstLineIndent());

            JSONObject option = allTheOptions.getJSONObject(i);
            String optionText = option.getString("value");
            boolean isChecked = checked.contains(option.getString("value"));
            // checkbox symbol in Wingdings 2 font
            String checkedText = isChecked ? "\u0052" : "\u00a3";
            XWPFRun rSymbol = newPara.createRun();
            rSymbol.setText(checkedText);
            rSymbol.setFontFamily("Wingdings 2");
            rSymbol.setFontSize(currPara.getRuns().get(0).getFontSizeAsDouble());
            String text = " " + optionText;
            XWPFRun r = newPara.createRun();
            r.setText(text);
            r.setFontFamily(fontFamily);
            r.setFontSize(fontSize);
        }
        return offset;
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
        String fontFamily = _getFontFamily(currPara);
        Double fontSize = _getFontSize(currPara);
        XmlCursor cursor = currPara.getCTP().newCursor();
        XWPFTable newTable = document.insertNewTbl(cursor);
        newTable.setTopBorder(XWPFBorderType.SINGLE , 5, 0, "000000");
        newTable.setBottomBorder(XWPFBorderType.SINGLE , 5, 0, "000000");
        newTable.setLeftBorder(XWPFBorderType.SINGLE , 5, 0, "000000");
        newTable.setRightBorder(XWPFBorderType.SINGLE , 5, 0, "000000");
        newTable.setInsideHBorder(XWPFBorderType.SINGLE , 5, 0, "000000");
        newTable.setInsideVBorder(XWPFBorderType.SINGLE , 5, 0, "000000");

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
            r.setFontFamily(fontFamily);
            r.setFontSize(fontSize);
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
                XWPFTableCell currCell = currRow.getCell(j);
                _writeCell(header.getJSONObject(j), innerArray.getString(j), currCell);

                for(XWPFParagraph p : currCell.getParagraphs()){
                    _setParagraphFontFamily(p, fontFamily);
                    _setParagraphFontSize(p, fontSize);
                }
            }
        }

        // 写表脚
        JSONArray footer = inputObj.getJSONObject("input_des").getJSONArray("rows");
        if (null == footer || footer.size() == 0){
            return 1;
        }
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
                        currRow.getCell(currCellIndex + colspan - 1).getCTTc().addNewTcPr().addNewHMerge()
                                .setVal(STMerge.CONTINUE);
                        mergedSize += colspan - 1;
                    }
                    break;
                case "sum":
                    try{
                        double sum = 0;
                        for (int j = 0; j < tableData.size(); j++) {
                            sum += Double.parseDouble(tableData.getJSONArray(j).getString(i + mergedSize));
                        }
                        if (sum == (long) sum) {
                            //如果结果是整数则部显示小数点
                            currCell.setText(String.format("%d", (long) sum));
                        } else {
                            currCell.setText(String.format("%s", sum));
                        }
                    }catch(Exception e){
                        currCell.setText("#N/A");
                    }
                    break;
                default:
                    currCell.setText(" ");
                    break;
            }

            XWPFRun r = currCell.getParagraphs().get(0).getRuns().get(0);
            r.setFontFamily(fontFamily);
            r.setFontSize(fontSize);
        }
        return 1;
    }

    /**
     * 根据列的类型，写表格
     * @param inputObj
     * @param formValue
     * @param cell
     * @return
     */
    private int _writeCell(JSONObject inputObj,String formValue, XWPFTableCell cell){
        int offset = 0;
        if (null == inputObj || null == formValue || null == cell) {
            return offset;
        }
        String inputType = inputObj.getString("input_type");
        if (null == inputType) {
            return offset;
        }
        switch (inputType) {
            case "file_list":
                XWPFParagraph paragraph = cell.getParagraphs().get(0);
                offset = _writeFileListCard(JSONArray.parse(formValue), paragraph);
                break;
            default:
                cell.setText(formValue);
                break;
        }
        return offset;
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
