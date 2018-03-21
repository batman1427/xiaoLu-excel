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

public class createXmxshk {

	public static void create(double[][] data,int years,int start) throws IOException {  
        String path = "E:/test/";  
        String fileName = "XXX项目销售回款情况";  
        String fileType = "xls";  
        writer(path, fileName, fileType,data,years,start);  
        System.out.println("创建销售回款情况成功");
    }  
	
	@SuppressWarnings("deprecation")
	public static void writer(String path, String fileName,String fileType,double[][] data,int years,int start) throws IOException {  
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 26)); 
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
        cell.setCellValue("XXX项目销售回款情况");
        
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
        Row row3 = sheet.createRow(3);
        sheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 1)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 3, 2, 2)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 5)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 6, 8)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 9, 11)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 3, 12, 12)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 13, 25)); 
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 13, 17)); 
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 18, 21)); 
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 22, 25)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 3, 26, 26)); 
        //项目	销售员	状态（在职/离职）	入职时间		总来访	11月来访
		

        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("项目");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("物业类型");
        Cell cell3 = row1.createCell(2);
        cell3.setCellStyle(style2);
        cell3.setCellValue("年份");
        Cell cell4 = row1.createCell(3);
        cell4.setCellStyle(style2);
        cell4.setCellValue("总认购");
        Cell cell5 = row1.createCell(6);
        cell5.setCellStyle(style2);
        cell5.setCellValue("总签约");
        Cell cell6 = row1.createCell(9);
        cell6.setCellStyle(style2);
        cell6.setCellValue("剩余未签约");									
        Cell cell7 = row1.createCell(12);
        cell7.setCellStyle(style2);
        cell7.setCellValue("库存套数");
        Cell cell8 = row1.createCell(13);
        cell8.setCellStyle(style2);
        cell8.setCellValue("按揭下款结算情况（注：该部分累加等于总签约部分）");
        Cell cell9 = row1.createCell(26);
        cell9.setCellStyle(style2);
        cell9.setCellValue("备注");
        
        
        Cell cell10 = row2.createCell(13);
        cell10.setCellStyle(style2);
        cell10.setCellValue("已下款已提交报告");
        Cell cell11 = row2.createCell(18);
        cell11.setCellStyle(style2);
        cell11.setCellValue("已下款未提交报告");
        Cell cell12 = row2.createCell(22);
        cell12.setCellStyle(style2);
        cell12.setCellValue("剩余未下款");
      
        Cell cell13 = row3.createCell(3);
        cell13.setCellStyle(style2);
        cell13.setCellValue("套数");
        Cell cell14 = row3.createCell(4);
        cell14.setCellStyle(style2);
        cell14.setCellValue("面积");
        Cell cell15 = row3.createCell(5);
        cell15.setCellStyle(style2);
        cell15.setCellValue("认购金额");
        Cell cell16 = row3.createCell(6);
        cell16.setCellStyle(style2);
        cell16.setCellValue("套数");
        Cell cell17 = row3.createCell(7);
        cell17.setCellStyle(style2);
        cell17.setCellValue("面积");
        Cell cell18 = row3.createCell(8);
        cell18.setCellStyle(style2);
        cell18.setCellValue("签约金额");
        Cell cell19 = row3.createCell(9);
        cell19.setCellStyle(style2);
        cell19.setCellValue("套数");
        Cell cell20 = row3.createCell(10);
        cell20.setCellStyle(style2);
        cell20.setCellValue("面积");
        Cell cell21 = row3.createCell(11);
        cell21.setCellStyle(style2);
        cell21.setCellValue("认购金额");
        Cell cell22 = row3.createCell(13);
        cell22.setCellStyle(style2);
        cell22.setCellValue("套数");
        Cell cell23 = row3.createCell(14);
        cell23.setCellStyle(style2);
        cell23.setCellValue("面积");
        Cell cell24 = row3.createCell(15);
        cell24.setCellStyle(style2);
        cell24.setCellValue("签约金额");
        Cell cell25 = row3.createCell(16);
        cell25.setCellStyle(style2);
        cell25.setCellValue("佣金预估");
        Cell cell26 = row3.createCell(17);
        cell26.setCellStyle(style2);
        cell26.setCellValue("备注（截止几月份报告）");
        Cell cell27 = row3.createCell(18);
        cell27.setCellStyle(style2);
        cell27.setCellValue("套数");
        Cell cell28 = row3.createCell(19);
        cell28.setCellStyle(style2);
        cell28.setCellValue("面积");
        Cell cell29 = row3.createCell(20);
        cell29.setCellStyle(style2);
        cell29.setCellValue("签约金额");
        Cell cell30 = row3.createCell(21);
        cell30.setCellStyle(style2);
        cell30.setCellValue("佣金预估");
        Cell cell31 = row3.createCell(22);
        cell31.setCellStyle(style2);
        cell31.setCellValue("套数");
        Cell cell32 = row3.createCell(23);
        cell32.setCellStyle(style2);
        cell32.setCellValue("面积");
        Cell cell33 = row3.createCell(24);
        cell33.setCellStyle(style2);
        cell33.setCellValue("签约金额");
        Cell cell34 = row3.createCell(25);
        cell34.setCellStyle(style2);
        cell34.setCellValue("佣金预估");
        
        for(int k=4;k<4*years+5;k++) {
        	 Row temp = sheet.createRow(k);
        	 if(k==4+years||k==5+years*2||k==6+years*3||k==7+years*3) {
        		 for(int i=0;i<23;i++) {
        			 Cell cella = temp.createCell(i+3);
        			 cella.setCellStyle(style2);
        			 cella.setCellValue(data[k-4][i]);
        		 }
        	 }else {
        		 for(int i=0;i<23;i++) {
        			 Cell cella = temp.createCell(i+3);
        			 cella.setCellValue(data[k-4][i]);
        		 }
        	 }
        }
        sheet.addMergedRegion(new CellRangeAddress(4, 15, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(4, 16, 26,26)); 
        sheet.addMergedRegion(new CellRangeAddress(4, 6,1, 1)); 
        sheet.addMergedRegion(new CellRangeAddress(8, 10,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(12, 14,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(7, 7,1, 2));
        sheet.addMergedRegion(new CellRangeAddress(11, 11,1, 2));
        sheet.addMergedRegion(new CellRangeAddress(15, 15,1, 2));
        sheet.addMergedRegion(new CellRangeAddress(16, 16,0, 2));
       	Row r1=sheet.getRow(4);
       	Cell c1 = r1.createCell(0);
       	c1.setCellStyle(style1);
        c1.setCellValue("XXX项目");
    	Cell c2 = r1.createCell(1);
       	c2.setCellStyle(style1);
        c2.setCellValue("住宅");
       
    	Row r2=sheet.getRow(4+years);
       	Cell c3 = r2.createCell(1);
       	c3.setCellStyle(style2);
        c3.setCellValue("合计");
        
        Row r3=sheet.getRow(5+years);
       	Cell c4 = r3.createCell(1);
       	c4.setCellStyle(style1);
        c4.setCellValue("商铺");
        
        Row r4=sheet.getRow(5+years*2);
       	Cell c5 = r4.createCell(1);
       	c5.setCellStyle(style2);
        c5.setCellValue("合计");
        
        Row r5=sheet.getRow(6+years*2);
       	Cell c6 = r5.createCell(1);
       	c6.setCellStyle(style1);
        c6.setCellValue("公寓");
        
        Row r6=sheet.getRow(6+years*3);
       	Cell c7 = r6.createCell(1);
       	c7.setCellStyle(style2);
        c7.setCellValue("合计");
        
        Row r7=sheet.getRow(7+years*3);
       	Cell c8 = r7.createCell(0);
       	c8.setCellStyle(style2);
        c8.setCellValue("总合计");
        
        for(int i=0;i<years;i++) {
        	Row r8=sheet.getRow(4+i);
        	Cell c9 = r8.createCell(2);
        	c9.setCellValue((start+i)+"年");
        }
        
        for(int i=0;i<years;i++) {
        	Row r8=sheet.getRow(5+i+years);
        	Cell c9 = r8.createCell(2);
        	c9.setCellValue((start+i)+"年");
        }
        
        for(int i=0;i<years;i++) {
        	Row r8=sheet.getRow(6+i+years*2);
        	Cell c9 = r8.createCell(2);
        	c9.setCellValue((start+i)+"年");
        }
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}
