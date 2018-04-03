package com.transform.rest;

import com.transform.file.tohtml.FileToHtml;
import com.transform.utils.FileReplacedUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import org.apache.commons.io.FileUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@Controller
@Api(tags = {"文件转换"})
@RequestMapping("/t")
public class TransformApi {

    FileToHtml htmlUtil = new FileToHtml();

    @PostMapping(value = "/transformToHtml", consumes = "multipart/*", headers = "content-type=multipart/form-data")
    @ApiOperation(value="文件转html")
    @ResponseBody
    public void toHtml(@ApiParam(value="file", required = true)MultipartFile file, HttpServletResponse response) throws IOException {
        String resMsg = "";

        String outTempFilePath = TransformApi.class.getClassLoader().getResource("").getPath();
        outTempFilePath = outTempFilePath.substring(0, outTempFilePath.indexOf("target")) + "html";
        File tempDir = new File(outTempFilePath);
        if(!tempDir.exists()){
            tempDir.mkdir();
        }

        File tempFile = new File(outTempFilePath + "/" + file.getOriginalFilename());

        FileUtils.copyInputStreamToFile(file.getInputStream(), tempFile);

        outTempFilePath += "/" +System.currentTimeMillis() + ".html";
        if(file.getOriginalFilename().endsWith(".xls") || file.getOriginalFilename().endsWith(".xlsx")) {
            resMsg = FileToHtml.excelToHtml(file.getInputStream(), outTempFilePath);
        }else if(file.getOriginalFilename().endsWith(".doc") || file.getOriginalFilename().endsWith(".docx")){
            resMsg = FileToHtml.wordToHtml(tempFile, outTempFilePath);
        }else{
            resMsg = "上传文件格式错误！！！";
        }

        String outfileName = file.getOriginalFilename().indexOf(".") > 0 ? file.getOriginalFilename().substring(0, file.getOriginalFilename().indexOf(".")) : file.getOriginalFilename();

        if(resMsg.startsWith("转换成功")){
            File f = new File(outTempFilePath);
            response.setCharacterEncoding("UTF-8");
            response.setContentType("text/html");
            response.setHeader("Content-Disposition", "attachment;fileName="+ outfileName +".html");

            try {
                FileInputStream inputStream = new FileInputStream(f);

                ServletOutputStream out = response.getOutputStream();

                int b = 0;
                byte[] buffer = new byte[2018];
                while (b != -1){
                    b = inputStream.read(buffer);
                    out.write(buffer,0,b);
                }

                out.close();
                out.flush();
                inputStream.close();

            } catch (IOException e) {
                e.printStackTrace();
            }finally {
                f.delete();
                tempFile.delete();
            }
        }else{
            throw new RuntimeException(resMsg);
        }


    }

    @PostMapping(value = "/placeholderReplace", consumes = "multipart/*", headers = "content-type=multipart/form-data")
    @ApiOperation(value="文档导出placeholder替换")
    @ResponseBody
    public void placeholderReplace(@ApiParam(value="file", required = true)MultipartFile file, HttpServletResponse response
            , @ApiParam(value="键值字符串, eg: id:123|name:张山", required = true)@RequestParam(value = "kvStr")String kvStr) throws IOException {

        kvStr = kvStr.replaceAll("：", ":");
        if(!file.getOriginalFilename().toUpperCase().endsWith(".XML")){
             throw  new RuntimeException("上传文件格式错误，需上传xml格式的文件！！！模板制作方法：打开word文件，另存为word xml格式。");
        }
        Map<String, Object> dataMap = new HashMap<String, Object>();
        String k = "";
        String kvArr[] = kvStr.split("\\|");
        for(int i = 0; i < kvArr.length; i ++){
            String str = kvArr[i].trim();
            if(str.indexOf(":") > 0 && !str.endsWith(":")){
                if(str.endsWith(":")) {
                    dataMap.put(str.substring(0, str.indexOf(":")), "");
                }else {
                    dataMap.put(str.substring(0, str.indexOf(":")), str.substring(str.indexOf(":") + 1));
                }
            }
        }



        File outFile;
        try {
            outFile = FileReplacedUtil.getPlacedWord(file.getInputStream(), dataMap);
            System.out.println(outFile.getName());
            System.out.println(outFile.getAbsolutePath());
            response.setCharacterEncoding("UTF-8");
            if(outFile.getName().toUpperCase().endsWith(".DOC")){
                response.setContentType("application/msword");
            }else if(outFile.getName().toUpperCase().endsWith(".XLS")){
                response.setContentType("application/x-xls");
            }

            response.setHeader("Content-Disposition", "attachment;fileName="+  file.getOriginalFilename().substring(0, file.getOriginalFilename().lastIndexOf("."))
                    + outFile.getName().substring(outFile.getName().lastIndexOf(".")));

            FileInputStream inputStream = new FileInputStream(outFile);
            ServletOutputStream out = response.getOutputStream();

            int b = 0;
            byte[] buffer = new byte[2018];
            while (b != -1){
                b = inputStream.read(buffer);
                out.write(buffer,0,b);
            }

            out.close();
            out.flush();
            inputStream.close();
            outFile.delete();
        }catch (Exception e){
            if(e.toString().indexOf("InvalidReferenceException") > 0){
                throw new RuntimeException("kv字符串格式不正确或字段不全！！！");
            }

        }

    }


}