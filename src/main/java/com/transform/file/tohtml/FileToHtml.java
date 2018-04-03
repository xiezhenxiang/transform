package com.transform.file.tohtml;

import java.io.File;
import java.io.InputStream;

public class FileToHtml {

    public static String excelToHtml(InputStream in, String outFilePath){
        return ExcelToHtml.read(in, outFilePath);
    }


    public static String wordToHtml(File file, String outFilePath){
        return WordToHtml.read(file, outFilePath);
    }

}
