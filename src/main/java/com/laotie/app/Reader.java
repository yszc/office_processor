package com.laotie.app;
import org.apache.poi.xwpf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Hello world!
 *
 */
public class Reader 
{
        public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream(new File("/workspaces/wordparser/template.docx"));
            XWPFDocument document = new XWPFDocument(fis);

            XWPFStyles style_sheet = document.getStyles();
            System.out.println(style_sheet.toString());


            // 编译正则表达式
            Pattern pattern = Pattern.compile("\\{([^}]+)\\}");
            for (IBodyElement element: document.getBodyElements()){
                System.out.println("========================================");
                if (element instanceof XWPFParagraph){
                    XWPFParagraph para = (XWPFParagraph) element;
                    String paraText = para.getText();
                    if(null == paraText || paraText.length()==0){
                        continue;
                    }
                    System.out.println("getText:"+para.getText());
                    System.out.println("getStyle:"+para.getStyle());
                    System.out.println("getStyleID:"+para.getStyleID());
                    String styleID = para.getStyleID();
                    if (null != styleID){
                        XWPFStyle style = style_sheet.getStyle(styleID);
                        System.out.println("style.getName:"+style.getName());
                        if (style.getName().indexOf("heading", 0)>=0){
                            
                        }
                    }

                    // 创建 Matcher 对象
                    Matcher matcher = pattern.matcher(paraText);
                    Boolean isMatched = false;
                    // 查找匹配
                    while (matcher.find()) {
                        isMatched = true;
                        String match = matcher.group(0); // 获取第一个捕获组的内容
                        paraText = paraText.replace(match, "这里是替换文本");
                    }
                    if (!isMatched){
                        continue;
                    }
                    int len = para.getRuns().size();
                    // 清除原段落中的内容块，只保留第0个
                    for (int i=len-1; i>=1; i--){
                        try{
                            para.removeRun(i);
                        }catch(Exception e){
                        }
                    }
                    // 设置新内容
                    XWPFRun run = para.getRuns().get(0);
                    run.setText(paraText, 0);
                    //run.setFontSize(30);
                    
                    System.out.println("getText:"+para.getText());
                
                }else if (element instanceof XWPFTable) {
                    System.out.println("Table found");
                    XWPFTable table = (XWPFTable) element;
                    for (XWPFTableRow row: table.getRows()) {
                        for (XWPFTableCell cell: row.getTableCells()){
                            System.err.println(cell.getText());
                        }
                    }
                }
            }
            FileOutputStream out = new FileOutputStream("replacement.docx");
            document.write(out);
            out.close();
            document.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
