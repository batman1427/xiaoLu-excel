package ExcelManage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createXmqd {
 
	public static void create(double[][] data,int year,int month,int day) throws IOException {  
        String path = "E:/test/";  
        String fileName = year+"年"+month+"月"+day+"日XXX项目渠道成交对比";  
        String fileType = "xls";  
        writer(path, fileName, fileType,data,year,month,day);  
        System.out.println("创建"+year+"年"+month+"月"+day+"日XXX项目渠道成交对比成功");
    }  
	
	@SuppressWarnings("deprecation")
	public static void writer(String path, String fileName,String fileType,double[][] data,int year,int month,int day) throws IOException {  
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 17)); 
        CellStyle style1 = wb.createCellStyle(); // 样式对象     
        style1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直      
        style1.setAlignment(CellStyle.ALIGN_CENTER);// 水平 
        style1.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());  
        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);  
        Font font1 = wb.createFont();  
        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);  
        font1.setFontName("宋体");  
        font1.setFontHeight((short) 280);  
        style1.setFont(font1); 
        Cell cell = row0.createCell(0);
        cell.setCellStyle(style1);
        cell.setCellValue(year+"年"+month+"月"+day+"日XXX项目渠道成交对比");
        
        CellStyle style2 = wb.createCellStyle(); // 样式对象     
        style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直      
        style2.setAlignment(CellStyle.ALIGN_CENTER);// 水平 
        style2.setWrapText(true);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框   
        Font font2 = wb.createFont();  
        font2.setFontName("宋体");  
        font2.setFontHeight((short) 280);  
        style2.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());  
        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);  
        font2.setFontHeight((short) 200); 
        style2.setFont(font2); 
        
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);      
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(1,2, 1, 1)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 5)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 9)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 10, 13)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 14, 17)); 

        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("序号");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("类型");
        Cell cell3 = row1.createCell(2);
        cell3.setCellStyle(style2);
        cell3.setCellValue("新润城");
        Cell cell4 = row1.createCell(6);
        cell4.setCellStyle(style2);
        cell4.setCellValue("中原");
        Cell cell5 = row1.createCell(10);
        cell5.setCellStyle(style2);
        cell5.setCellValue("自销");
        Cell cell6 = row1.createCell(14);
        cell6.setCellStyle(style2);
        cell6.setCellValue("同策");									
       
        Cell cell7 = row2.createCell(2);       
        cell7.setCellStyle(style2);
        cell7.setCellValue("本日认购套数");               
        Cell cell8 = row2.createCell(3);
        cell8.setCellStyle(style2);
        cell8.setCellValue("本日认购额");
        Cell cell9 = row2.createCell(4);
        cell9.setCellStyle(style2);
        cell9.setCellValue("累计认购套数");  
        Cell cell10 = row2.createCell(5);
        cell10.setCellStyle(style2);
        cell10.setCellValue("累计认购额");
        Cell cell11 = row2.createCell(6);
        cell11.setCellStyle(style2);
        cell11.setCellValue("本日认购套数");
        Cell cell12 = row2.createCell(7);
        cell12.setCellStyle(style2);
        cell12.setCellValue("本日认购额");      
        Cell cell13 = row2.createCell(8);
        cell13.setCellStyle(style2);
        cell13.setCellValue("累计认购套数");
        Cell cell14 = row2.createCell(9);
        cell14.setCellStyle(style2);
        cell14.setCellValue("累计认购额");
        Cell cell15 = row2.createCell(10);
        cell15.setCellStyle(style2);
        cell15.setCellValue("本日认购套数");
        Cell cell16 = row2.createCell(11);
        cell16.setCellStyle(style2);
        cell16.setCellValue("本日认购额");
        Cell cell17 = row2.createCell(12);
        cell17.setCellStyle(style2);
        cell17.setCellValue("累计认购套数");
        Cell cell18 = row2.createCell(13);
        cell18.setCellStyle(style2);
        cell18.setCellValue("累计认购额");
        Cell cell19 = row2.createCell(14);
        cell19.setCellStyle(style2);
        cell19.setCellValue("本日认购套数");
        Cell cell20 = row2.createCell(15);
        cell20.setCellStyle(style2);
        cell20.setCellValue("本日认购额");
        Cell cell21 = row2.createCell(16);
        cell21.setCellStyle(style2);
        cell21.setCellValue("累计认购套数");
        Cell cell22 = row2.createCell(17);
        cell22.setCellStyle(style2);
        cell22.setCellValue("累计认购额");
       
         for(int k=3;k<10;k++) {
        	 Row temp = sheet.createRow(k);	
        	 for(int i=0;i<16;i++) {
       			 Cell celltemp = temp.createCell(i+2);
       			 celltemp.setCellStyle(style2);
       			 celltemp.setCellValue(data[k-3][i]);
        	 }
       	 
        }
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 0, 1)); 
       
        Row r3=sheet.getRow(3);
        Cell cell30 = r3.createCell(0);
        cell30.setCellStyle(style2);
		cell30.setCellValue("1");
		Cell cell31 = r3.createCell(1);
		cell31.setCellStyle(style2);
	    cell31.setCellValue("自然来访成交");
		
		Row r4=sheet.getRow(4);
		Cell cell40 = r4.createCell(0);
		cell40.setCellStyle(style2);
		cell40.setCellValue("2");
		Cell cell41 = r4.createCell(1);
		cell41.setCellStyle(style2);
		cell41.setCellValue("中原中介带访成交");
		
		Row r5=sheet.getRow(5);
		Cell cell50 = r5.createCell(0);
		cell50.setCellStyle(style2);
		cell50.setCellValue("3");
		Cell cell51 = r5.createCell(1);
		cell51.setCellStyle(style2);
		cell51.setCellValue("小鹿中介带访成交");
		
		Row r6=sheet.getRow(6);
		Cell cell60 = r6.createCell(0);
		cell60.setCellStyle(style2);
		cell60.setCellValue("4");
		Cell cell61 = r6.createCell(1);
		cell61.setCellStyle(style2);
		cell61.setCellValue("老带新成交");
		
		Row r7=sheet.getRow(7);
		Cell cell70 = r7.createCell(0);
		cell70.setCellStyle(style2);
		cell70.setCellValue("5");
		Cell cell71 = r7.createCell(1);
		cell71.setCellStyle(style2);
		cell71.setCellValue("全民营销推荐");
		
		Row r8=sheet.getRow(8);
		Cell cell80 = r8.createCell(0);
		cell80.setCellStyle(style2);
		cell80.setCellValue("6");
		Cell cell81 = r8.createCell(1);
		cell81.setCellStyle(style2);
		cell81.setCellValue("其他中介");
		
		Row r9=sheet.getRow(9);
		Cell cell90 = r9.createCell(0);
		cell90.setCellStyle(style2);
		cell90.setCellValue("合计");
		
		
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}

