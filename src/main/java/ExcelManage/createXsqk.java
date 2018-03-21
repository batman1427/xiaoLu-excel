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

public class createXsqk {
    
	public static void create(double[][] data,int year,int month,int day) throws IOException {  
        String path = "E:/test/";  
        String fileName = year+"年"+month+"月"+day+"日XXX项目销售情况";  
        String fileType = "xls";  
        writer(path, fileName, fileType,data,year,month,day);  
        System.out.println("创建"+year+"年"+month+"月"+day+"日XXX项目销售情况成功");
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 14)); 
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
        cell.setCellValue(year+"年"+month+"月"+day+"日XXX项目销售情况");
        
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
        sheet.addMergedRegion(new CellRangeAddress(1,2, 1, 1)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 5)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 8)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 9, 11)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 12, 14)); 

        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("序号");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("类别");
        Cell cell3 = row1.createCell(2);
        cell3.setCellStyle(style2);
        cell3.setCellValue("公司");
        Cell cell4 = row1.createCell(3);
        cell4.setCellStyle(style2);
        cell4.setCellValue("新润城");
        Cell cell5 = row1.createCell(6);
        cell5.setCellStyle(style2);
        cell5.setCellValue("中原");
        Cell cell6 = row1.createCell(9);
        cell6.setCellStyle(style2);
        cell6.setCellValue("甲方");									
        Cell cell7 = row1.createCell(12);       
        cell7.setCellStyle(style2);
        cell7.setCellValue("同策");
               
        Cell cell8 = row2.createCell(3);
        cell8.setCellStyle(style2);
        cell8.setCellValue("套数");
        Cell cell9 = row2.createCell(4);
        cell9.setCellStyle(style2);
        cell9.setCellValue("面积");  
        Cell cell10 = row2.createCell(5);
        cell10.setCellStyle(style2);
        cell10.setCellValue("金额");
        Cell cell11 = row2.createCell(6);
        cell11.setCellStyle(style2);
        cell11.setCellValue("套数");
        Cell cell12 = row2.createCell(7);
        cell12.setCellStyle(style2);
        cell12.setCellValue("面积");      
        Cell cell13 = row2.createCell(8);
        cell13.setCellStyle(style2);
        cell13.setCellValue("金额");
        Cell cell14 = row2.createCell(9);
        cell14.setCellStyle(style2);
        cell14.setCellValue("套数");
        Cell cell15 = row2.createCell(10);
        cell15.setCellStyle(style2);
        cell15.setCellValue("面积");
        Cell cell16 = row2.createCell(11);
        cell16.setCellStyle(style2);
        cell16.setCellValue("金额");
        Cell cell17 = row2.createCell(12);
        cell17.setCellStyle(style2);
        cell17.setCellValue("套数");
        Cell cell18 = row2.createCell(13);
        cell18.setCellStyle(style2);
        cell18.setCellValue("面积");
        Cell cell19 = row2.createCell(14);
        cell19.setCellStyle(style2);
        cell19.setCellValue("金额");

        for(int k=3;k<23;k++) {
       	 Row temp = sheet.createRow(k);	
       	 for(int i=0;i<12;i++) {
       			 Cell celltemp = temp.createCell(i+3);
       			 celltemp.setCellValue(data[k-3][i]);
       	 }
       	 
       }
        Row aaa = sheet.createRow(23);	
        Row bbb = sheet.createRow(24);	
        Row ccc = sheet.createRow(25);	
        Row ddd = sheet.createRow(26);	
        sheet.addMergedRegion(new CellRangeAddress(3, 5, 0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(3, 5, 1,1)); 
        sheet.addMergedRegion(new CellRangeAddress(6, 9,0, 0)); 
        sheet.addMergedRegion(new CellRangeAddress(6, 9,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(10, 13,0, 0));
        sheet.addMergedRegion(new CellRangeAddress(10, 13,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(14, 17,0, 0));
        sheet.addMergedRegion(new CellRangeAddress(14, 17,1,1));
        sheet.addMergedRegion(new CellRangeAddress(19, 22,0, 0));
        sheet.addMergedRegion(new CellRangeAddress(19, 22,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(23, 26,0, 0));
        sheet.addMergedRegion(new CellRangeAddress(23, 26,1, 1));
        sheet.addMergedRegion(new CellRangeAddress(23, 23,3, 14));
        sheet.addMergedRegion(new CellRangeAddress(24, 24,3, 14));
        sheet.addMergedRegion(new CellRangeAddress(25, 25,3, 14));
        sheet.addMergedRegion(new CellRangeAddress(26, 26,3, 14));
        Row r3=sheet.getRow(3);
        Cell cell30 = r3.createCell(0);
		cell30.setCellValue("1");
		Cell cell31 = r3.createCell(1);
	    cell31.setCellValue("认筹");
		Cell cell32 = r3.createCell(2);
		cell32.setCellValue("今日");
		
		Row r4=sheet.getRow(4);
		Cell cell42 = r4.createCell(2);
		cell42.setCellValue("本月");
		
		Row r5=sheet.getRow(5);
		Cell cell52 = r5.createCell(2);
		cell52.setCellValue("累计");
		
		Row r6=sheet.getRow(6);
		Cell cell60 = r6.createCell(0);
		cell60.setCellValue("2");
		Cell cell61 = r6.createCell(1);
		cell61.setCellValue("认购（住宅）");
		Cell cell62 = r6.createCell(2);
		cell62.setCellValue("今日");
		
		Row r7=sheet.getRow(7);
		Cell cell72 = r7.createCell(2);
		cell72.setCellValue("本月");
		
		Row r8=sheet.getRow(8);
		Cell cell82 = r8.createCell(2);
		cell82.setCellValue("本年认购");
		
		Row r9=sheet.getRow(9);
		Cell cell92 = r9.createCell(2);
		cell92.setCellValue("累计认购");
		
		Row r10=sheet.getRow(10);
		Cell cell100 = r10.createCell(0);
		cell100.setCellValue("2");
		Cell cell101 = r10.createCell(1);
		cell101.setCellValue("认购（商业）");
		Cell cell102 = r10.createCell(2);
		cell102.setCellValue("今日");
		
		Row r11=sheet.getRow(11);
		Cell cell112 = r11.createCell(2);
		cell112.setCellValue("本月");
		
		Row r12=sheet.getRow(12);
		Cell cell122 = r12.createCell(2);
		cell122.setCellValue("本年认购");
		
		Row r13=sheet.getRow(13);
		Cell cell132 = r13.createCell(2);
		cell132.setCellValue("累计认购");
		
		Row r14=sheet.getRow(14);
		Cell cell140 = r14.createCell(0);
		cell140.setCellValue("3");
		Cell cell141 = r14.createCell(1);
		cell141.setCellValue("签约");
		Cell cell142 = r14.createCell(2);
		cell142.setCellValue("今日签约");
		
		Row r15=sheet.getRow(15);
		Cell cell152 = r15.createCell(2);
		cell152.setCellValue("本月签约");
		
		Row r16=sheet.getRow(16);
		Cell cell162 = r16.createCell(2);
		cell162.setCellValue("本年签约");
		
		Row r17=sheet.getRow(17);
		Cell cell172 = r17.createCell(2);
		cell172.setCellValue("累计总签约");
		
		Row r18=sheet.getRow(18);
		Cell cell180 = r18.createCell(0);
		cell180.setCellValue("4");
		Cell cell181 = r18.createCell(1);
		cell181.setCellValue("未签约");
		Cell cell182 = r18.createCell(2);
		cell182.setCellValue("未签约");
		
		Row r19=sheet.getRow(19);
		Cell cell190 = r19.createCell(0);
		cell190.setCellValue("5");
		Cell cell191 = r19.createCell(1);
		cell191.setCellValue("下款及回款情况");
		Cell cell192 = r19.createCell(2);
		cell192.setCellValue("已签约已下款");
		
		Row r20=sheet.getRow(20);
		Cell cell202 = r20.createCell(2);
		cell202.setCellValue("已下款已结佣");
		
		Row r21=sheet.getRow(21);
		Cell cell212 = r21.createCell(2);
		cell212.setCellValue("已签约未下款");
		
		Row r22=sheet.getRow(22);
		Cell cell222 = r22.createCell(2);
		cell222.setCellValue("已下款未结佣");
		
		Row r23=sheet.getRow(23);
		Cell cell230 = r23.createCell(0);
		cell230.setCellValue("6");
		Cell cell231 = r23.createCell(1);
		cell231.setCellValue("其他事务");
		Cell cell232 = r23.createCell(2);
		cell232.setCellValue("今日案场内部工作摘要");
		Cell cell233 = r23.createCell(3);
		cell233.setCellValue("1.正常盘客 2.call客成果检核 3.外展地点人员分布");
		
		Row r24=sheet.getRow(24);
		Cell cell242 = r24.createCell(2);
		cell242.setCellValue("今日与甲方对接事宜");	
		Cell cell243 = r24.createCell(3);
		cell243.setCellValue("无");
		
		Row r25=sheet.getRow(25);
		Cell cell252 = r25.createCell(2);
		cell252.setCellValue("今日营销活动");	
		Cell cell253 = r25.createCell(3);
		cell253.setCellValue("无");
		
		Row r26=sheet.getRow(26);
		Cell cell262 = r26.createCell(2);
		cell262.setCellValue("其他事项");	
		Cell cell263 = r26.createCell(3);
		cell263.setCellValue("无");
       	
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}

