package com.laotie.app;

import static com.laotie.app.WordParser.parseInput;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONObject;

public class WordWriter extends WordParser {

    public WordWriter(String filePath) throws IOException {
        super(filePath);
    }

    /**
     * 设置段落内容
     * TODO: check if changed
     * @param paragraph
     * @param content
     * @return
     */
    private XWPFParagraph setParagraphContent(XWPFParagraph paragraph, String content) {
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
        return paragraph;
    }

    public void writeTemplate(String jsonString, String saveFile) throws IOException {
        JSONObject jsonObj = JSON.parseObject(jsonString);
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                String content = paragraph.getText();
                for (String json : extractJson(content)) {
                    JSONObject inputObj = JSON.parseObject(json);
                    if (null == inputObj.get("var_name")) {
                        continue;
                    }
                    // TODO: content replace
                    // Object inputInfo = jsonObj.get(inputObj.get("var_name"));
                    // content = content.replace(json, inputInfo.toString());
                }
                paragraph = setParagraphContent(paragraph, content);
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
        } catch (IOException e){
            e.printStackTrace();
        }
    }

}
