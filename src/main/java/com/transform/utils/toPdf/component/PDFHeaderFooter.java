package com.transform.utils.toPdf.component;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import com.transform.utils.toPdf.component.builder.HeaderFooterBuilder;

import java.io.IOException;

/**
 * Created by fgm on 2017/4/22.
 * 页眉页脚定制工具
 */
public class PDFHeaderFooter implements HeaderFooterBuilder {
    /**
     * @param writer   PDF编写类
     * @param document PDF文档对象
     * @param data     业务数据
     * @param font     字体设置
     * @param template PDF模板
     * @description PDF页脚设置类
     */
    public void writeFooter(PdfWriter writer,
                            Document document,
                            Object data,
                            Font font,
                            PdfTemplate template) {
        if (data == null) {
            return;
        }
        int pageS = writer.getPageNumber();
        int currentPage = pageS;
        if (currentPage <= 0) {
            return;
        }

        Phrase footer2 = new Phrase(currentPage + " /", font);

        PdfContentByte cb = writer.getDirectContent();

        ColumnText.showTextAligned(
                cb,
                Element.ALIGN_RIGHT,
                footer2,
                (document.right() - 10),
                document.top() + 18, 0);

        //设置模板位置
        cb.addTemplate(template, document.right() - 10, document.top() + 18);

    }

    /**
     * @param writer   PDF编写类
     * @param document PDF文档对象
     * @param data     业务数据
     * @param font     字体设置
     * @param template PDF模板
     * @description PDF页头设置类
     */
    public void writeHeader(PdfWriter writer,
                            Document document,
                            Object data,
                            Font font,
                            PdfTemplate template) throws IOException, DocumentException {
        String imgPath = PDFHeaderFooter.class.getClassLoader().getResource("").getPath() +"templates/pdfHead.png";
        imgPath = imgPath.replace("test-classes", "classes");
        Image imghead = Image.getInstance(imgPath);

        imghead.setAlignment(Image.ALIGN_CENTER);
        imghead.scaleToFit(100, 180);

        imghead.setAbsolutePosition(50, 10);

        PdfContentByte cbhead = writer.getDirectContent();
        PdfTemplate tphead = cbhead.createTemplate(300, 300);
        tphead.addImage(imghead);
        cbhead.addTemplate(tphead, -15, 800);
        ColumnText.showTextAligned(
                writer.getDirectContent(),
                Element.ALIGN_LEFT,
                new Phrase(  "", font),
                document.left(),
                document.top() + 18, 0);


    }


    /**
     * @param writer   PDF编写类
     * @param document PDF文档对象
     * @param data     业务数据
     * @description 页头、页眉设置的模板替换类
     */
    public String getReplaceOfTemplate(PdfWriter writer, Document document, Object data) {
        int total = writer.getPageNumber() - 1;
        return " " + total;
    }

}
