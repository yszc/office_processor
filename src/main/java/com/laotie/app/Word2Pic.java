package com.laotie.app;
import com.spire.doc.*;
import com.spire.doc.documents.ImageType;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;

public class Word2Pic {
    
    public static void main(String[] args) throws Exception {
        //创建Document实例
        Document doc = new Document();
        //加载Word文档
        doc.loadFromFile("docs/jsoncase.docx");

        //转换到图片并设置图片的分辨率
        BufferedImage[] images = doc.saveToImages(0, doc.getPageCount(), ImageType.Bitmap, 500, 500);

        int i = 0;
        for (BufferedImage image : images) {
            //保存为.png文件格式
            File file = new File( "doc/" + String.format(("Img-%d.png"), i));
            ImageIO.write(image, "PNG", file);
            i++;
        }
    }
    
}
