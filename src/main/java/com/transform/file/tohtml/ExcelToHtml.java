package com.transform.file.tohtml;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @功能描述 POI 读取 Excel 转 HTML 支持 03xls 和 07xlsx 版本  包含样式
 * @author xzx
 * @创建时间 2018/03/19
 */
public class ExcelToHtml {


    public static String read(InputStream in, String outFilePath){

        String message = "";
        String htmlExcel = null;
        try {
            Workbook wb = WorkbookFactory.create(in);//此WorkbookFactory在POI-3.10版本中使用需要添加dom4j
            if (wb instanceof XSSFWorkbook) {
                XSSFWorkbook xWb = (XSSFWorkbook) wb;
                htmlExcel = ExcelToHtml.getExcelInfo(xWb,true);
            }else if(wb instanceof HSSFWorkbook){
                HSSFWorkbook hWb = ( HSSFWorkbook) wb;
                htmlExcel = ExcelToHtml.getExcelInfo(hWb,true);
            }
            //System.out.println(htmlExcel);
            FileOutputStream fos = null;
            BufferedWriter bw = null;
            try {
                //输出文件路径


                File file = new File(outFilePath);
                fos = new FileOutputStream(file);
                bw = new BufferedWriter(new OutputStreamWriter(fos));
                bw.write(htmlExcel);
            } catch (FileNotFoundException fnfe) {
                message = "系统找不到输出路径！！！";
            } catch (IOException ioe) {
                ioe.printStackTrace();
            } finally {
                try {
                    if (bw != null)
                        bw.close();
                    if (fos != null)
                        fos.close();
                } catch (IOException ie) {
                }
            }

        } catch (IllegalArgumentException e) {
            message = "文件格式有误！！！";
        } catch (InvalidFormatException e) {
            message = "文件格式有误！！！";
        } catch (FileNotFoundException e) {
            message = "输入文件路径不存在！！！";
        } catch (IOException e) {
            e.printStackTrace();
        } finally{
            try {
                in.close();
            } catch (Exception e) {
                message = "输入文件路径不存在！！！";
            }
        }
        if(message.equals(""))
            message = "转换成功！输出路径为：" + outFilePath;
        return message;
    }

    public static boolean isHeadOrTail = false;
    public static  double widthFactor = 0.85;//自定义表格自适应宽度系数

    /**
     * 程序入口方法
     * @param filePath 文件的路径
     * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
     * @return <table>...</table> 字符串
     */
    public String readExcelToHtml(String filePath , boolean isWithStyle){

        InputStream is = null;
        String htmlExcel = null;
        try {
            File sourcefile = new File(filePath);
            is = new FileInputStream(sourcefile);
            Workbook wb = WorkbookFactory.create(is);
            if (wb instanceof XSSFWorkbook) {
                XSSFWorkbook xWb = (XSSFWorkbook) wb;
                htmlExcel = ExcelToHtml.getExcelInfo(xWb,isWithStyle);
            }else if(wb instanceof HSSFWorkbook){
                HSSFWorkbook hWb = (HSSFWorkbook) wb;
                htmlExcel = ExcelToHtml.getExcelInfo(hWb,isWithStyle);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return htmlExcel;
    }



    public static String getExcelInfo(Workbook wb,boolean isWithStyle){


        String sbArr[] = new String[100];

        for(int i = 0; i < 500; i ++) {
            StringBuffer sb = new StringBuffer();
            try{
                Sheet sheet = wb.getSheetAt(i);//获取第一个Sheet的内容
            }catch (Exception e){
                break;
            }
            Sheet sheet = wb.getSheetAt(i);

            //System.out.println(getValidColumnNum(sheet));
            int lastRowNum = sheet.getLastRowNum();
            Map<String, String> map[] = getRowSpanColSpanMap(sheet);
            sb.append("<table style='border-collapse:collapse;' width='1000pxpx'>");
            Row row = null;        //兼容
            Cell cell = null;    //兼容
            int totalColspan = getValidColumnNum(sheet);
            double totalWidth = 0.0;
            for(int j = 0; j < totalColspan; j ++){
                totalWidth += sheet.getColumnWidthInPixels(j);
            }
            for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
                isHeadOrTail = false;

                row = sheet.getRow(rowNum);

                if (row == null || isEmptyRow(sheet, row)) {
                    sb.append("<tr><td>&nbsp;</td></tr>");
                    continue;
                }
                sb.append("<tr>");
                int lastColNum = row.getLastCellNum();
                int endcol = 0;int colSpan = 0;
                for (int colNum = 0; colNum < lastColNum; colNum++) {

                    if (colNum >= totalColspan)
                        break;
                    cell = row.getCell(colNum);

                    if (cell == null) {    //特殊情况 空白的单元格会返回null
                        sb.append("<td></td>");
                        continue;
                    }

                    String stringValue = getCellValue(cell);


                    if (map[0].containsKey(rowNum + "," + colNum)) {
                        String pointString = map[0].get(rowNum + "," + colNum);
                        map[0].remove(rowNum + "," + colNum);
                        int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                        int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                        int rowSpan = bottomeRow - rowNum + 1;
                        colSpan = bottomeCol - colNum + 1;
                        if(rowNum == sheet.getFirstRowNum() && colSpan == totalColspan) {
                            sb = new StringBuffer(sb.toString().substring(0,sb.length() - 4));
                            sb.append("<caption ");
                            isHeadOrTail = true;
                        }else
                            sb.append("<td rowspan= '" + rowSpan + "' colspan= '" + colSpan + "' ");
                        endcol = colNum + colSpan - 1;

                    } else if (map[1].containsKey(rowNum + "," + colNum)) {
                        map[1].remove(rowNum + "," + colNum);
                        continue;
                    } else {
                        sb.append("<td ");
                        endcol = colNum;

                    }

                    //判断是否需要样式
                    if (isWithStyle) {
                        dealExcelStyle(wb, sheet, cell, sb, colNum, endcol);//处理单元格样式
                    }

                    sb.append(">");
                    if (stringValue == null || "".equals(stringValue.trim())) {
                        sb.append("<input type=\"text\" value=\"\"  style=\"width:98%;height: 25px; font-size: 20px;border:none;\"/>");
                    } else {
                        // 将ascii码为160的空格转换为html下的空格（&nbsp;）
                        sb.append(stringValue.replace(String.valueOf((char) 160), "&nbsp;"));
                    }
                    if(isHeadOrTail)
                        sb.append("</caption>");
                    else
                        sb.append("</td>");
                }
                if( !isHeadOrTail)
                    sb.append("</tr>");
            }

            sb.append("</table><br/><br/>");
            sb = new StringBuffer(sb.toString().replace("1000pxpx", (totalWidth / widthFactor) + "px"));
            sbArr[i] = sb.toString();
        }
        String rs = "<!DOCTYPE html><html><head><meta charset=\"UTF-8\"></head><body>";
        for(int i = 0; i < sbArr.length; i ++) {
            if(sbArr[i] == null || sbArr[i].equals(""))
                break;
            rs += sbArr[i];
        }
        rs += "</body></html>";
        return rs;
    }

    private static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {

        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            // System.out.println(topRow + "," + topCol + "," + bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = { map0, map1 };
        return map;
    }


    /**
     * 获取表格单元格Cell内容
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {

        String result = new String();
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = DateUtil
                            .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case Cell.CELL_TYPE_STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case Cell.CELL_TYPE_BLANK:
                result = "";
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 处理表格样式
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     */
    private static void dealExcelStyle(Workbook wb,Sheet sheet,Cell cell,StringBuffer sb, int colStart, int colEnd){
        int totalColspan = getValidColumnNum(sheet);
        float totalWidth = 0;
        for(int i = 0; i < totalColspan; i ++){
            totalWidth += sheet.getColumnWidthInPixels(i);
        }



        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignment();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getVerticalAlignment();
            sb.append("valign='"+ convertVerticalAlignToHtml(verticalAlignment)+ "' ");//单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {

                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                short boldWeight = xf.getBoldweight();
                sb.append("style='");
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + xf.getFontHeight() / 2 + "%;"); // 字体大小
                float columnWidth = 0;
                for(int i = colStart; i <= colEnd; i ++){
                    columnWidth += sheet.getColumnWidthInPixels(Integer.parseInt(i + ""));
                }
                String tt = (double)columnWidth / (double)totalWidth * 100 + "%";
                sb.append("width:" + tt +"; ");

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
                }

                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                //System.out.println("************************************");
                //System.out.println("BackgroundColorColor: "+cellStyle.getFillBackgroundColorColor());
                //System.out.println("ForegroundColor: "+cellStyle.getFillForegroundColor());//0
                //System.out.println("BackgroundColorColor: "+cellStyle.getFillBackgroundColorColor());
                //System.out.println("ForegroundColorColor: "+cellStyle.getFillForegroundColorColor());
                //String bgColorStr = bgColor.getARGBHex();
                //System.out.println("bgColorStr: "+bgColorStr);
                if (bgColor != null && !"".equals(bgColor)) {
                    sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); // 背景颜色
                }

                sb.append(getBorderStyle(0,cellStyle.getBorderTop(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
                sb.append(getBorderStyle(1,cellStyle.getBorderRight(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
                sb.append(getBorderStyle(2,cellStyle.getBorderBottom(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
                sb.append(getBorderStyle(3,cellStyle.getBorderLeft(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));

            }else if(wb instanceof HSSFWorkbook){

                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                short boldWeight = hf.getBoldweight();
                short fontColor = hf.getColor();
                sb.append("style='");
                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 字体颜色
                }
                double columnWidth = 0;
                for(int i = colStart; i <= colEnd; i ++){
                    columnWidth += sheet.getColumnWidthInPixels(i);
                }
                String tt = (double)columnWidth / (double)totalWidth * 100 + "%";

                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
                }
                if(!isHeadOrTail) {
                    sb.append("width:" + tt + "; ");
                    boolean flag = true;

                    if(!(getBorderStyle(palette, 0, cellStyle.getBorderTop(), cellStyle.getTopBorderColor()).indexOf("#d0d7e5") > 0 && getBorderStyle(palette, 1, cellStyle.getBorderRight(), cellStyle.getRightBorderColor()).indexOf("#d0d7e5") > 0
                        && getBorderStyle(palette, 3, cellStyle.getBorderLeft(), cellStyle.getLeftBorderColor()).indexOf("#d0d7e5") > 0 && getBorderStyle(palette, 2, cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor()).indexOf("#d0d7e5") > 0)) {
                            sb.append(getBorderStyle(palette, 0, cellStyle.getBorderTop(), cellStyle.getTopBorderColor()));
                            sb.append(getBorderStyle(palette, 1, cellStyle.getBorderRight(), cellStyle.getRightBorderColor()));
                            sb.append(getBorderStyle(palette, 3, cellStyle.getBorderLeft(), cellStyle.getLeftBorderColor()));
                            sb.append(getBorderStyle(palette, 2, cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor()));
                    }
                }
            }

            sb.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式
     * @param alignment
     * @return
     */
    private static String convertAlignToHtml(short alignment) {

        String align = "left";
        switch (alignment) {
            case CellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case CellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case CellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 单元格中内容的垂直排列方式
     * @param verticalAlignment
     * @return
     */
    private static String convertVerticalAlignToHtml(short verticalAlignment) {

        String valign = "middle";
        switch (verticalAlignment) {
            case CellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case CellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case CellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    private static String convertToStardColor(HSSFColor hc) {

        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }

        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    static String[] bordesr={"border-top:","border-right:","border-bottom:","border-left:"};
    static String[] borderStyles={"solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid","solid","solid","solid","solid"};

    private static  String getBorderStyle(  HSSFPalette palette ,int b,short s, short t){

        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        String borderColorStr = convertToStardColor( palette.getColor(t));
        borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr;
        return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";

    }

    private static  String getBorderStyle(int b,short s, XSSFColor xc){

        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();//t.getARGBHex();
            borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr.substring(2);
            return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
        }

        return "";
    }

    private static int getValidColumnNum(Sheet sheet){
        int num = 0;
        int f[] = new int[300];
        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if(row != null && row.getPhysicalNumberOfCells() > 0){
                f [row.getLastCellNum()] += 1;
            }
        }
        for(int i = 1; i < 300; i ++){
            if(f[i] > num)
                num = i;
        }

        return num;
    }

    private  static boolean isEmptyRow(Sheet sheet, Row row){

        int totalColumnNum = getValidColumnNum(sheet);


        for(int i = 0; i < totalColumnNum; i ++){
            try {

                if(row.getCell(i) != null && getCellValue(row.getCell(i)) != null && !getCellValue(row.getCell(i)).trim().equals("")){
                    return false;
                }
            }catch (Exception e){
                continue;
            }
        }

        return true;

    }
}