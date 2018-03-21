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

public class createCjmx {

	public static void create(String[][] data) throws IOException {  
        String path = "E:/test/";  
        String fileName = "新润城老带新全民营销成交明细";  
        String fileType = "xls";  
        writer(path, fileName, fileType,data);  
        System.out.println("创建新润城老带新全民营销成交明细成功");
    }  
	
	@SuppressWarnings("deprecation")
	public static void writer(String path, String fileName,String fileType,String[][] data) throws IOException {  
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9)); 
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
        cell.setCellValue("新润城老带新全民营销成交明细");
        
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
        
        CellStyle style3 = wb.createCellStyle(); // 样式对象     
        style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直      
        style3.setAlignment(CellStyle.ALIGN_CENTER);// 水平 
        
        Row row1 = sheet.createRow(1);		

        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("序号");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("成交日期");
        Cell cell3 = row1.createCell(2);
        cell3.setCellStyle(style2);
        cell3.setCellValue("推荐人");
        Cell cell4 = row1.createCell(3);
        cell4.setCellStyle(style2);
        cell4.setCellValue("推荐人联系电话");
        Cell cell5 = row1.createCell(4);
        cell5.setCellStyle(style2);
        cell5.setCellValue("客户姓名");
        Cell cell6 = row1.createCell(5);
        cell6.setCellStyle(style2);
        cell6.setCellValue("成交房号");									
        Cell cell7 = row1.createCell(6);
        cell7.setCellStyle(style2);
        cell7.setCellValue("销售额");
        Cell cell8 = row1.createCell(7);
        cell8.setCellStyle(style2);
        cell8.setCellValue("置业顾问");
        Cell cell9 = row1.createCell(8);
        cell9.setCellStyle(style2);
        cell9.setCellValue("推荐人奖励");       
        Cell cell10 = row1.createCell(9);
        cell10.setCellStyle(style2);
        cell10.setCellValue("备注");
       
        
        for(int k=2;k<2+data.length;k++) {
        	 Row temp = sheet.createRow(k);    
        	 Cell cellte = temp.createCell(0);
        	 cellte.setCellStyle(style3);
        	 cellte.setCellValue(k-2);
        	 for(int i=0;i<9;i++) {
        			Cell celltemp = temp.createCell(i+1);
        			celltemp.setCellValue(data[k-2][i]);
        	 }
        }
       
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}

