package kratos.demo;

import java.awt.image.BufferedImage;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import javax.imageio.ImageIO;

import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;

/**
 * Hello world!
 * 
 * @author test
 */
public class PdfConvert {

    public static void main(String[] args) {
        String path = "./";
        PdfConvert app = new PdfConvert();
        app.pdfConvert(path);
    }

    public void pdfConvert(String current) {
        File file = new File(current);
        List<String> pathList = getAllFilePath(file);
        // 4. 循环转换pdf
        pathList.forEach(filePath -> {
            String pdfL = ".pdf";
            if (!filePath.toLowerCase().contains(pdfL)) {
                System.out.println("pdf转图片失败, 传入文件不是pdf格式, filePath = " + filePath);
            } else {
                try {
                    File newFile = new File(filePath);
                    System.out.println("开始pdf转图片, filePath = " + filePath + ", size = " + newFile.length());
                    PDDocument doc = PDDocument.load(newFile, MemoryUsageSetting.setupTempFileOnly());
                    PDFRenderer renderer = new PDFRenderer(doc);
                    int pageCount = doc.getNumberOfPages();
                    BufferedImage image;
                    for (int index = 0; index < pageCount; index++) {
                        String tempImgPath = newFile.getName() + "_" + (index + 1) + ".jpg";
                        PDRectangle cropBox = doc.getPage(index).getCropBox();
                        float height = cropBox.getHeight();
                        // 固定的dpi值太大，会导致有一些大图片内存溢出，太小会导致小图片不清晰，所以动态设置一下
                        image = renderer.renderImageWithDPI(index, 72 * 1440 / height);
                        File tempImgFile = new File(tempImgPath);
                        ImageIO.write(image, "jpg", tempImgFile);
                        System.out.println("pdf转图片成功, fileName = " + tempImgPath + ", index = " + index);
                    }
                    doc.close();
                } catch (Exception e) {
                    System.out.println("pdf转图片失败, pdf解析失败, filePath = " + filePath + ", e = " + e);
                }
            }
        });
    }

    private List<String> getAllFilePath(File file) {
        boolean directory = file.isDirectory();
        List<String> filePathList = new ArrayList<>();
        if (directory) {
            for (File listFile : Objects.requireNonNull(file.listFiles())) {
                filePathList.addAll(getAllFilePath(listFile));
            }
        } else {
            filePathList.add(file.getAbsolutePath());
        }
        return filePathList;
    }

}
