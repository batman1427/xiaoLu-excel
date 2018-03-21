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

public class createTjbpm {

	public static void create(ArrayList<String> name,double[][] data,int year,int month) throws IOException {  
        String path = "E:/test/";  
        String fileName = year+"年"+month+"月销售排名";  
        String fileType = "xls";  
        writer(path, fileName, fileType,name,data,year,month);  
        System.out.println(year+"年"+month+"月销售排名成功");
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 24)); 
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
        cell.setCellValue(year+"年XXX项目"+month+"月度销售排名");
        
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
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 5)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 9)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 10, 14)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 15, 19)); 
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 20, 24)); 
        //项目	销售员	状态（在职/离职）	入职时间		总来访	11月来访
		

        Cell cell1 = row1.createCell(0);
        cell1.setCellStyle(style2);
        cell1.setCellValue("项目");
        Cell cell2 = row1.createCell(1);
        cell2.setCellStyle(style2);
        cell2.setCellValue("销售员	");
        Cell cell3 = row1.createCell(2);
        cell3.setCellStyle(style2);
        cell3.setCellValue("状态（在职/离职）");
        Cell cell4 = row1.createCell(3);
        cell4.setCellStyle(style2);
        cell4.setCellValue("入职时间");
        Cell cell5 = row1.createCell(4);
        cell5.setCellStyle(style2);
        cell5.setCellValue("总来访");
        Cell cell6 = row1.createCell(5);
        cell6.setCellStyle(style2);
        cell6.setCellValue(month+"月来访");									
        Cell cell7 = row1.createCell(6);
        cell7.setCellStyle(style2);
        cell7.setCellValue(month+"月认购");
        Cell cell8 = row1.createCell(10);
        cell8.setCellStyle(style2);
        cell8.setCellValue("累计认购");
        Cell cell9 = row1.createCell(15);
        cell9.setCellStyle(style2);
        cell9.setCellValue(month+"月签约");
        Cell cell10 = row1.createCell(20);
        cell10.setCellStyle(style2);
        cell10.setCellValue("累计签约");
        
        Cell cell11 = row2.createCell(6);
        cell11.setCellStyle(style2);
        cell11.setCellValue("套数");
        Cell cell12 = row2.createCell(7);
        cell12.setCellStyle(style2);
        cell12.setCellValue("面积");
        Cell cell13 = row2.createCell(8);
        cell13.setCellStyle(style2);
        cell13.setCellValue("销售额");
        Cell cell14 = row2.createCell(9);
        cell14.setCellStyle(style2);
        cell14.setCellValue("成交率");
        Cell cell15 = row2.createCell(10);
        cell15.setCellStyle(style2);
        cell15.setCellValue("套数");
        Cell cell16 = row2.createCell(11);
        cell16.setCellStyle(style2);
        cell16.setCellValue("面积");
        Cell cell17 = row2.createCell(12);
        cell17.setCellStyle(style2);
        cell17.setCellValue("销售额");
        Cell cell18 = row2.createCell(13);
        cell18.setCellStyle(style2);
        cell18.setCellValue("成交率");
        Cell cell19 = row2.createCell(14);
        cell19.setCellStyle(style2);
        cell19.setCellValue("排名");
        Cell cell20 = row2.createCell(15);
        cell20.setCellStyle(style2);
        cell20.setCellValue("套数");
        Cell cell21 = row2.createCell(16);
        cell21.setCellStyle(style2);
        cell21.setCellValue("面积");
        Cell cell22 = row2.createCell(17);
        cell22.setCellStyle(style2);
        cell22.setCellValue("签约额");
        Cell cell23 = row2.createCell(18);
        cell23.setCellStyle(style2);
        cell23.setCellValue("转签率");
        Cell cell24 = row2.createCell(19);
        cell24.setCellStyle(style2);
        cell24.setCellValue("排名");
        Cell cell25 = row2.createCell(20);
        cell25.setCellStyle(style2);
        cell25.setCellValue("套数");
        Cell cell26 = row2.createCell(21);
        cell26.setCellStyle(style2);
        cell26.setCellValue("面积");
        Cell cell27 = row2.createCell(22);
        cell27.setCellStyle(style2);
        cell27.setCellValue("签约额");
        Cell cell28 = row2.createCell(23);
        cell28.setCellStyle(style2);
        cell28.setCellValue("转签率");
        Cell cell29 = row2.createCell(24);
        cell29.setCellStyle(style2);
        cell29.setCellValue("排名");
        
        int start=3;
        int end=3+name.size();
        for(int k=start;k<=end;k++) {
        	 Row temp = sheet.createRow(k);
        	 if(k!=end) {
        		 Cell cella = temp.createCell(1);
        		 cella.setCellValue(name.get(k-start));
        		 Cell cellb = temp.createCell(2);
        		 cellb.setCellValue("在职");
        		 Cell cellc = temp.createCell(3);
        		 cellc.setCellValue("2017/00/00");
        	 }
             Cell celld = temp.createCell(4);
             celld.setCellValue(data[k-start][0]);
             Cell celle = temp.createCell(5);
             celle.setCellValue(data[k-start][1]);
             Cell cellf = temp.createCell(6);
             cellf.setCellValue(data[k-start][2]);
             Cell cellg = temp.createCell(7);
             cellg.setCellValue(data[k-start][3]);
             Cell cellh = temp.createCell(8);
             cellh.setCellValue(data[k-start][4]);
             Cell celli = temp.createCell(9);
             celli.setCellValue(data[k-start][5]);
             Cell cellj = temp.createCell(10);
             cellj.setCellValue(data[k-start][6]);
             Cell cellk = temp.createCell(11);
             cellk.setCellValue(data[k-start][7]);
             Cell celll = temp.createCell(12);
             celll.setCellValue(data[k-start][8]);
             Cell cellm = temp.createCell(13);
             cellm.setCellValue(data[k-start][9]);
             Cell celln = temp.createCell(14);
             celln.setCellValue(data[k-start][10]);
             Cell cello = temp.createCell(15);
             cello.setCellValue(data[k-start][11]);
             Cell cellp = temp.createCell(16);
             cellp.setCellValue(data[k-start][12]);
             Cell cellq = temp.createCell(17);
             cellq.setCellValue(data[k-start][13]);
             Cell cellr = temp.createCell(18);
             cellr.setCellValue(data[k-start][14]);
             Cell cells = temp.createCell(19);
             cells.setCellValue(data[k-start][15]);
             Cell cellt = temp.createCell(20);
             cellt.setCellValue(data[k-start][16]);
             Cell cellu = temp.createCell(21);
             cellu.setCellValue(data[k-start][17]);
             Cell cellv = temp.createCell(22);
             cellv.setCellValue(data[k-start][18]);
             Cell cellw = temp.createCell(23);
             cellw.setCellValue(data[k-start][19]);
             Cell cellx = temp.createCell(24);
             cellx.setCellValue(data[k-start][20]);
             if(k==end) {
            	 Cell cellend = temp.createCell(0);
                 cellend.setCellValue("合计：");
             }
        }
       	sheet.addMergedRegion(new CellRangeAddress(3, 2+name.size(), 0, 0)); 
       	Row project=sheet.getRow(3);
       	Cell cellend = project.createCell(0);
       	cellend.setCellStyle(style2);
        cellend.setCellValue("XXX项目");
        
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}
