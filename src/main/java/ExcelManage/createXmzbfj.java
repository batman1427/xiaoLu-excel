package ExcelManage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createXmzbfj {

	
	public static void create(ArrayList<String> name,double[][] data,int year,int month) throws IOException {  
        String path = "E:/test/";  
        String fileName = year+"年"+month+"月项目指标分解";  
        String fileType = "xls";  
        writer(path, fileName, fileType,name,data,year,month);  
        System.out.println(year+"年"+month+"月项目指标分解创建成功");
    }  
	
	@SuppressWarnings("deprecation")
	public static void writer(String path, String fileName,String fileType,ArrayList<String> name,double[][] data,int year,int month) throws IOException {  
        Workbook wb = null; 
        String excelPath = path+File.separator+fileName+"."+fileType;
        File file = new File(excelPath);
        Sheet sheet =null;
        //创建工作文档对象   
        if (!file.exists()) {
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook();
                
            } else if(fileType.equals("xlsx")) {
                
                wb = new XSSFWorkbook();
            } else {
                System.out.println("文件格式不正确");
            }
            //创建sheet对象   
            sheet = (Sheet) wb.createSheet("sheet1");  
            OutputStream outputStream = new FileOutputStream(excelPath);
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
            
        } else {
            if (fileType.equals("xls")) {  
                wb = new HSSFWorkbook();  
                
            } else if(fileType.equals("xlsx")) { 
                wb = new XSSFWorkbook();  
                
            } else {  
            	 System.out.println("文件格式不正确");
            }  
        }
        //创建sheet对象   
        if (sheet==null) {
            sheet = (Sheet) wb.createSheet("sheet1");  
        }
        
        //第一行 
        Row row0 = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 16)); 
        CellStyle style1 = wb.createCellStyle(); // 样式对象     
        style1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直      
        style1.setAlignment(CellStyle.ALIGN_CENTER);// 水平  
        Font font1 = wb.createFont();  
        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);  
        font1.setFontName("宋体");  
        font1.setFontHeight((short) 280);  
        style1.setFont(font1); 
        Cell cell = row0.createCell(0);
        cell.setCellStyle(style1);
        cell.setCellValue("XXX项目"+year+"年"+month+"月业绩目标制定情况汇总表");
        
        CellStyle style2 = wb.createCellStyle(); // 样式对象     
        style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直      
        style2.setAlignment(CellStyle.ALIGN_CENTER);// 水平 
        style2.setWrapText(true);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框   
        Font font2 = wb.createFont();  
        font2.setBoldweight(Font.BOLDWEIGHT_BOLD);  
        font2.setFontName("宋体");  
        font2.setFontHeight((short) 280);  
        style2.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());  
        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);  
        font2.setFontHeight((short) 200); 
        style2.setFont(font2); 
        
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 9)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 10, 15)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 16, 16)); 
        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("姓名");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("岗位");
        Cell cell3 = row1.createCell(2);
        cell3.setCellStyle(style2);
        cell3.setCellValue("甲方指标分解");
        Cell cell4 = row1.createCell(3);
        cell4.setCellStyle(style2);
        cell4.setCellValue("项目本月指标");
        Cell cell5 = row1.createCell(10);
        cell5.setCellStyle(style2);
        cell5.setCellValue("项目本月实际完成情况");
        Cell cell6 = row1.createCell(16);
        cell6.setCellStyle(style2);
        cell6.setCellValue("备注");									

        Cell cell7 = row2.createCell(3);
        cell7.setCellStyle(style2);
        cell7.setCellValue("认购套数");
        Cell cell8 = row2.createCell(4);
        cell8.setCellStyle(style2);
        cell8.setCellValue("认购金额（万元）");
        Cell cell9 = row2.createCell(5);
        cell9.setCellStyle(style2);
        cell9.setCellValue("签约套数");
        Cell cell10 = row2.createCell(6);
        cell10.setCellStyle(style2);
        cell10.setCellValue("签约金额（万元）");
        Cell cell11 = row2.createCell(7);
        cell11.setCellStyle(style2);
        cell11.setCellValue("佣金回款套数");
        Cell cell12 = row2.createCell(8);
        cell12.setCellStyle(style2);
        cell12.setCellValue("佣金回款（万元）");
        Cell cell13 = row2.createCell(9);
        cell13.setCellStyle(style2);
        cell13.setCellValue("其它回款（万元）");
        Cell cell14 = row2.createCell(10);
        cell14.setCellStyle(style2);
        cell14.setCellValue("认购套数");
        Cell cell15 = row2.createCell(11);
        cell15.setCellStyle(style2);
        cell15.setCellValue("认购金额（万元）");
        Cell cell16 = row2.createCell(12);
        cell16.setCellStyle(style2);
        cell16.setCellValue("签约套数");
        Cell cell17 = row2.createCell(13);
        cell17.setCellStyle(style2);
        cell17.setCellValue("签约金额（万元）");
        Cell cell18 = row2.createCell(14);
        cell18.setCellStyle(style2);
        cell18.setCellValue("佣金回款（万元）");
        Cell cell19 = row2.createCell(15);
        cell19.setCellStyle(style2);
        cell19.setCellValue("其它回款（万元）");
        
        int start=3;
        int end=2+name.size();
        for(int k=start;k<=end;k++) {
        	 Row temp = sheet.createRow(k);
        	 Cell cella = temp.createCell(0);
             cella.setCellValue(name.get(k-start));
             Cell cellb = temp.createCell(1);
             cellb.setCellValue("置业顾问");
             Cell cellc = temp.createCell(10);
             cellc.setCellValue(data[k-start][0]);
             Cell celld = temp.createCell(11);
             celld.setCellValue(data[k-start][1]);
             Cell celle = temp.createCell(12);
             celle.setCellValue(data[k-start][2]);
             Cell cellf = temp.createCell(13);
             cellf.setCellValue(data[k-start][3]);
             Cell cellg = temp.createCell(14);
             cellg.setCellValue(data[k-start][4]);
        }
        
        Row count = sheet.createRow(end+1);
        Cell cell20 = count.createCell(0);
        cell20.setCellValue("合计：");
        Cell cell21 = count.createCell(3);
        cell21.setCellValue("0");
        Cell cell22 = count.createCell(4);
        cell22.setCellValue("0");
        Cell cell23 = count.createCell(5);
        cell23.setCellValue("0");
        Cell cell24 = count.createCell(6);
        cell24.setCellValue("0");
        Cell cell25 = count.createCell(10);
        cell25.setCellValue("/");
        Cell cell26 = count.createCell(11);
        cell26.setCellValue("/");
        Cell cell27 = count.createCell(12);
        cell27.setCellValue("/");
        Cell cell28 = count.createCell(13);
        cell28.setCellValue("/");
    
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}
