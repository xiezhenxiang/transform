package com.transform.sensor.common;

import com.transform.utils.toPdf.component.PDFKit;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.util.HashMap;
import java.util.Map;

@RunWith(SpringJUnit4ClassRunner.class)
@SpringBootTest
public class AppTest {

    @Test
    public void test(){}

    @Test
    public void toPdf(){
        String resMsg = "";
        Map<String, Object> dataMap= new HashMap<String, Object>();
        dataMap.put("title", "Just For Test!");
        resMsg = PDFKit.createPDF(dataMap,"test.ftl");
        System.out.println(resMsg);
    }
}
