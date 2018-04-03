package com.transform.utils.toPdf.component.builder;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.IOException;

/**
 * Created by fgm on 2017/4/15.
 */
public interface HeaderFooterBuilder {


    /**
     * @description 写页眉
     */
    void writeHeader(PdfWriter writer, Document document, Object data, Font font, PdfTemplate template) throws IOException, DocumentException;
    /**
     * @description 写页脚
     */
    void writeFooter(PdfWriter writer, Document document, Object data, Font font, PdfTemplate template);

    /**
     * @description 关闭文档前,获取替换页眉页脚处设置模板的文本
     */
    String  getReplaceOfTemplate(PdfWriter writer, Document document, Object data);

}
