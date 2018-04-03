package com.transform.file.tohtml;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

public class WordToHtml {

    public static String read(File file, String outFilePath) {

        String message = "";


        try {
            if (file.getName().endsWith(".docx") || file.getName().endsWith(".DOCX")) {
                docx(file, "", outFilePath);
            } else {
                dox(file, "", outFilePath);
            }
        }catch (Exception e){
            if(e.toString().startsWith("java.io.FileNotFoundException")){
                    message = "系统找不到输出路径！！！";
            }else if(e.toString().indexOf("NotOLE2FileException") >= 0){
                message = "输入文件格式异常！！！";
            }
        }
        if(message.equals(""))
            message = "转换成功！输出路径为：" + outFilePath;
        return message;
    }

    /**
     * 转换docx
     * @param fileName
     * @param htmlName
     * @throws Exception
     */
    public static void docx(File f ,String fileName,String htmlName) throws Exception{

        // ) 加载word文档生成 XWPFDocument对象
        InputStream in = new FileInputStream(f);
        XWPFDocument document = new XWPFDocument(in);
        // ) 解析 XHTML配置 (这里设置IURIResolver来设置图片存放的目录)

        XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(f));
        options.setExtractor(new FileImageExtractor(f));
        options.setIgnoreStylesIfUnused(false);
        options.setFragment(true);
        // ) 将 XWPFDocument转换成XHTML
        OutputStream out = new FileOutputStream(new File(htmlName));

        XHTMLConverter.getInstance().convert(document, out, options);
        System.out.println(out);
    }
    /**
     * 转换doc
     * @param fileName
     * @param htmlName
     * @throws Exception
     */
    public static void dox(File file ,String fileName,String htmlName) throws Exception{

        InputStream input = new FileInputStream(file);
        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        //解析word文档
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();

        File htmlFile = new File(htmlName);
        OutputStream outStream = new FileOutputStream(htmlFile);

        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);

        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");

        serializer.transform(domSource, streamResult);
        outStream.close();
    }
}
