package com.transform.utils;

import freemarker.template.Configuration;
import freemarker.template.Template;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.util.Map;

public class FileReplacedUtil {

    private static Configuration configuration;
    private static String encoding;
    public FileReplacedUtil(){}

    public FileReplacedUtil(String encoding) {
        this.encoding = encoding;
        configuration = new Configuration(Configuration.VERSION_2_3_22);
        configuration.setDefaultEncoding(encoding);
    }

    public static  String  getTempPath(){

        return FileReplacedUtil.class.getClassLoader().getResource("").getPath();

    }

    public static Template getTemplate(String path) throws Exception {
        return configuration.getTemplate(path);
    }

    public String getImageStr(String image) throws IOException {
        InputStream is = new FileInputStream(image);
        BASE64Encoder encoder = new BASE64Encoder();
        byte[] data = new byte[is.available()];
        is.read(data); is.close();
        return encoder.encode(data);
    }


    public static void exportDoc(String doc, String templateName, Map<String, Object> dataMap) throws Exception {
        configuration.setDirectoryForTemplateLoading(new File(getTempPath()));
        Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(doc), encoding));
        getTemplate(templateName).process(dataMap, writer);
    }

    public static File getPlacedWord (InputStream in, Map<String, Object> dataMap) throws Exception {

        FileReplacedUtil maker = new FileReplacedUtil("UTF-8");

        SAXReader reader = new SAXReader();
        // 调用读方法，得到Document
        Document doc = reader.read(in);

        String xmlStr = doc.asXML();

        //解决占位符分离问题
        int pos = 0;
        while (pos < xmlStr.length() && xmlStr.indexOf("$", pos) > 0){
            int index = xmlStr.indexOf("$", pos);
            String str = xmlStr.substring(index, xmlStr.length()).replaceAll(" ", "").replaceAll("\r\n", "")
                    .replaceAll("\t", "").replaceAll("\n", "");

            int end = str.indexOf("}");
            int nextIndex = str.indexOf("$", 1);
            boolean isSureEnd = end > 1 && (nextIndex < 0 || (nextIndex > 0 && end < nextIndex));
            //去掉制表符标签
            if(str.charAt(1) == '{' && isSureEnd){
                int tt = xmlStr.indexOf("}", index) + 1;
                String oldStr = xmlStr.substring(index, tt);
                String newStr = str.substring(0, end + 1);
                while (newStr.indexOf("<") > 0 && newStr.indexOf(">") > 0){
                    newStr = newStr.replace(newStr.substring(newStr.indexOf("<"), newStr.indexOf(">") + 1), "");
                }
                xmlStr = xmlStr.replace(oldStr, newStr);
                String key = newStr.replace("{", "").replace("}", "").replace("$", "");
                if(dataMap.get(key) == null){
                    dataMap.put(key, "");
                }
            }

            pos =index + 1;
        }

        String tempFileName = "";
        if(xmlStr.indexOf("Word.Document") > 0){
            tempFileName = "temp.doc";
        }else if(xmlStr.indexOf("Excel.Sheet") > 0){
            tempFileName = "temp.xls";
        }
        //System.out.println(xmlStr);
        Document document = DocumentHelper.parseText(xmlStr);
        //写入emp1.xml文件
        OutputFormat outputFormat = new OutputFormat();
        outputFormat.setEncoding("UTF-8");

        OutputStream outputStream = new FileOutputStream(getTempPath() + "temp.xml");
        XMLWriter xmlWriter = new XMLWriter(outputStream,outputFormat);

        xmlWriter.write(document);
        xmlWriter.flush();
        xmlWriter.close();

        maker.exportDoc(getTempPath() + tempFileName, "temp.xml", dataMap);

        return new File(getTempPath() + tempFileName);

       /* ActiveXComponent _app = new ActiveXComponent("Excel.Application");
        _app.setProperty("Visible", Variant.VT_FALSE);
        Dispatch documents = _app.getProperty("Workbooks").toDispatch();
        Dispatch doc = Dispatch.call(documents, "Open", "C:/Users/XZX/Desktop/NewDoc.xlsx").toDispatch();
        Dispatch.call(doc, "SaveAs", "C:/Users/XZX/Desktop/tt.xlsx");
        Dispatch.call(doc, "Close", Variant.VT_FALSE);
        _app.invoke("Quit", new Variant[] {});
        ComThread.Release();*/


        /*ActiveXComponent _app = new ActiveXComponent("Word.Application");
    	_app.setProperty("Visible", Variant.VT_FALSE);

    	Dispatch documents = _app.getProperty("Documents").toDispatch();

    	// 打开FreeMarker生成的Word文档
    	Dispatch doc = Dispatch.call(documents, "Open", "C:/Users/XZX/Desktop/t5.doc", Variant.VT_FALSE, Variant.VT_TRUE).toDispatch();
    	// 另存为新的Word文档
    	Dispatch.call(doc, "SaveAs", "C:/Users/XZX/Desktop/NewDoc.doc", Variant.VT_FALSE, Variant.VT_TRUE);

    	Dispatch.call(doc, "Close", Variant.VT_FALSE);
    	_app.invoke("Quit", new Variant[] {});
    	ComThread.Release();*/
    }



}
