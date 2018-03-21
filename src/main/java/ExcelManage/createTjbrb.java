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

public class createTjbrb {
	public static void create(double[][] data,int year,int month,int day,int type,ArrayList<String> zjtype) throws IOException {  
        String path = "E:/test/";  
        String fileName = year+"年"+month+"月"+day+"日XXX项目蓄客情况";  
        String fileType = "xls";  
        writer(path, fileName, fileType,data,year,month,day,type,zjtype);  
        System.out.println("创建"+year+"年"+month+"月"+day+"日XXX项目蓄客情况成功");
    }  
	
	@SuppressWarnings("deprecation")
	public static void writer(String path, String fileName,String fileType,double[][] data,int year,int month,int day,int type,ArrayList<String> zjtype) throws IOException {  
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12)); 
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
        cell.setCellValue(year+"年"+month+"月"+day+"日XXX项目蓄客情况");
        
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
        style3.setWrapText(true);
        style3.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
        style3.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
        style3.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
        style3.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框   
        Font font3 = wb.createFont();  
        //font2.setBoldweight(Font.BOLDWEIGHT_BOLD);  
        font3.setFontName("宋体");  
        font3.setFontHeight((short) 280);  
        style3.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());  
        style3.setFillPattern(CellStyle.SOLID_FOREGROUND);  
        font3.setFontHeight((short) 200); 
        style3.setFont(font3); 
        
        Row row1 = sheet.createRow(1);
        Row row2 = sheet.createRow(2);
        
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 1,2)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 4)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 5, 6)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 7, 8)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 9, 10)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 11, 12));
		

     
        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("序号");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("类别");
        Cell cell3 = row1.createCell(3);
        cell3.setCellStyle(style2);
        cell3.setCellValue("今日（组）");
        Cell cell4 = row1.createCell(5);
        cell4.setCellStyle(style2);
        cell4.setCellValue("本月（组）");
        Cell cell5 = row1.createCell(7);
        cell5.setCellStyle(style2);
        cell5.setCellValue("累计（组）");
        Cell cell6 = row1.createCell(9);
        cell6.setCellStyle(style2);
        cell6.setCellValue("认筹情况（组）");									
        Cell cell7 = row1.createCell(11);
        cell7.setCellStyle(style2);
        cell7.setCellValue("成交情况 （组）");
       
        
        Cell cell8 = row2.createCell(3);
        cell8.setCellStyle(style2);
        cell8.setCellValue("合计");
        Cell cell9 = row2.createCell(4);
        cell9.setCellStyle(style2);
        cell9.setCellValue("有效");
        Cell cell10 = row2.createCell(5);
        cell10.setCellStyle(style2);
        cell10.setCellValue("合计");
        Cell cell11 = row2.createCell(6);
        cell11.setCellStyle(style2);
        cell11.setCellValue("有效");
        Cell cell12 = row2.createCell(7);
        cell12.setCellStyle(style2);
        cell12.setCellValue("总累计");
        Cell cell13 = row2.createCell(8);
        cell13.setCellStyle(style2);
        cell13.setCellValue("有效客户");
        Cell cell14 = row2.createCell(9);
        cell14.setCellStyle(style2);
        cell14.setCellValue("今日");
        Cell cell15 = row2.createCell(10);
        cell15.setCellStyle(style2);
        cell15.setCellValue("累计");
        Cell cell16 = row2.createCell(11);
        cell16.setCellStyle(style2);
        cell16.setCellValue("今日");
        Cell cell17 = row2.createCell(12);
        cell17.setCellStyle(style2);
        cell17.setCellValue("累计");
       
        
        for(int k=3;k<data.length+3;k++) {
        	 Row temp = sheet.createRow(k);	
        	 for(int i=0;i<10;i++) {
        			 Cell celltemp = temp.createCell(i+3);
        			 celltemp.setCellStyle(style3);
        			 celltemp.setCellValue(data[k-3][i]);
        	 }
        	 
        }
        if(type>1) {
        sheet.addMergedRegion(new CellRangeAddress(3, 2+type, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(3, 2+type, 1,2)); 
        }
        sheet.addMergedRegion(new CellRangeAddress(3+type,3+type,1, 2)); 
        sheet.addMergedRegion(new CellRangeAddress(4+type, 5+type,0, 0));
        sheet.addMergedRegion(new CellRangeAddress(4+type, 5+type,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(6+type,8+type,0, 0));
        sheet.addMergedRegion(new CellRangeAddress(6+type,8+type,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(9+type, 9+type,1, 2));
        sheet.addMergedRegion(new CellRangeAddress(10+type, 10+type,1, 2));
        sheet.addMergedRegion(new CellRangeAddress(11+type,11+type,1, 2));
        sheet.addMergedRegion(new CellRangeAddress(12+type, 12+type,0, 2));
        for(int i=0;i<type;i++) {
        	if(i==0) {
        		Row r1=sheet.getRow(i+3);
               	Cell c1 = r1.createCell(0);
               	c1.setCellStyle(style3);
                c1.setCellValue("1");
                Cell c2 = r1.createCell(1);
               	c2.setCellStyle(style3);
                c2.setCellValue("中介带访");
        	}
        	Row r2=sheet.getRow(i+3);
           	Cell c3 = r2.createCell(2);
           	c3.setCellStyle(style3);
            c3.setCellValue(zjtype.get(i));
        }
        Row a=sheet.getRow(type+3);
        Cell cella = a.createCell(0);
        cella.setCellStyle(style3);
		cella.setCellValue("2");
		Cell cellb = a.createCell(1);
	    cellb.setCellStyle(style3);
	    cellb.setCellValue("外展情况");
		
	    Row c=sheet.getRow(type+4);
        Cell cellc = c.createCell(0);
        cellc.setCellStyle(style3);
		cellc.setCellValue("3");
		Cell celld = c.createCell(1);
	    celld.setCellStyle(style3);
	    celld.setCellValue("外拓情况");
	    Cell celle = c.createCell(2);
	    celle.setCellStyle(style3);
	    celle.setCellValue("留电");
	    Row f=sheet.getRow(type+5);
        Cell cellf = f.createCell(2);
        cellf.setCellStyle(style3);
		cellf.setCellValue("拉访");
       
		Row g=sheet.getRow(type+6);
        Cell cellg = g.createCell(0);
        cellg.setCellStyle(style3);
		cellg.setCellValue("4");
		Cell cellh = g.createCell(1);
	    cellh.setCellStyle(style3);
	    cellh.setCellValue("CALL客");
	    Cell celli = g.createCell(2);
	    celli.setCellStyle(style3);
	    celli.setCellValue("电商后台推送数据");
	    Row j=sheet.getRow(type+7);
        Cell cellj = j.createCell(2);
        cellj.setCellStyle(style3);
		cellj.setCellValue("call客户资源");
		Row k=sheet.getRow(type+8);
	    Cell cellk = k.createCell(2);
	    cellk.setCellStyle(style3);
	    cellk.setCellValue("小计");
	    
	    Row l=sheet.getRow(type+9);
        Cell celll = l.createCell(0);
        celll.setCellStyle(style3);
		celll.setCellValue("5");
		Cell cellm = l.createCell(1);
	    cellm.setCellStyle(style3);
	    cellm.setCellValue("来电情况");
	    
	    Row n=sheet.getRow(type+10);
        Cell celln = n.createCell(0);
        celln.setCellStyle(style3);
		celln.setCellValue("6");
		Cell cello = n.createCell(1);
	    cello.setCellStyle(style3);
	    cello.setCellValue("新客户自然来访情况");
	    
	    Row p=sheet.getRow(type+11);
        Cell cellp = p.createCell(0);
        cellp.setCellStyle(style3);
		cellp.setCellValue("7");
		Cell cellq = p.createCell(1);
	    cellq.setCellStyle(style3);
	    cellq.setCellValue("老客户来访");
	    
	    Row r=sheet.getRow(type+12);
        Cell cellr = r.createCell(0);
        cellr.setCellStyle(style3);
		cellr.setCellValue("合计");
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}
