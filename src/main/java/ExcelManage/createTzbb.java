package ExcelManage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createTzbb {
	public static void create(String year,double[][] data) throws IOException {  
        String path = "E:/test/";  
        String fileName = year+"年台账报表";  
        String fileType = "xls";  
        writer(path, fileName, fileType,data);  
        System.out.println(year+"台账报表创建成功");
    }  
    
    public static void writer(String path, String fileName,String fileType,double[][] data) throws IOException {  
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
        
        //添加表头  
        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);
        Cell cell = row1.createCell(0);
        cell.setCellValue("项目：");    //创建第一行    
        
   
        CellStyle style = wb.createCellStyle(); // 样式对象      
        // 设置单元格的背景颜色为淡蓝色  
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);  
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直      
        style.setAlignment(CellStyle.ALIGN_CENTER);// 水平   
       
        Row row2 = sheet.createRow(2);   
        cell = row2.createCell(0);  
        cell.setCellStyle(style);
        cell.setCellValue("项目");  
        for(int i = 1;i <13 ;i++){  
            cell = row2.createCell(i);  
            cell.setCellStyle(style);
            cell.setCellValue(i+"月");  
        }  
        //row.setHeight((short) 540); 

        Row row3 = sheet.createRow(3);   
        cell = row3.createCell(0);  
        cell.setCellValue("认购套数");  
        for(int i = 1;i <13 ;i++){  
            cell = row3.createCell(i);  
            cell.setCellValue(data[0][i-1]);  
        } 
        
        Row row4 = sheet.createRow(4);   
        cell = row4.createCell(0);  
        cell.setCellValue("认购面积");  
        for(int i = 1;i <13 ;i++){  
            cell = row4.createCell(i);  
            cell.setCellValue(data[1][i-1]);  
        }  
        
        Row row5 = sheet.createRow(5);   
        cell = row5.createCell(0);  
        cell.setCellValue("认购金额");  
        for(int i = 1;i <13 ;i++){  
            cell = row5.createCell(i);  
            cell.setCellValue(data[2][i-1]);  
        }  
        
        Row row6 = sheet.createRow(6);   
       
        
        Row row7 = sheet.createRow(7);   
        cell = row7.createCell(0);  
        cell.setCellValue("签约套数");  
        for(int i = 1;i <13 ;i++){  
            cell = row7.createCell(i);  
            cell.setCellValue(data[3][i-1]);  
        }  
        
        Row row8 = sheet.createRow(8);   
        cell = row8.createCell(0);  
        cell.setCellValue("签约面积");  
        for(int i = 1;i <13 ;i++){  
            cell = row8.createCell(i);  
            cell.setCellValue(data[4][i-1]);  
        }  
        
        Row row9 = sheet.createRow(9);   
        cell = row9.createCell(0);  
        cell.setCellValue("合同金额");  
        for(int i = 1;i <13 ;i++){  
            cell = row9.createCell(i);  
            cell.setCellValue(data[5][i-1]);  
        }  
        
        Row row10 = sheet.createRow(10);   
        cell = row10.createCell(0);  
        cell.setCellValue("合计佣金");  
        for(int i = 1;i <13 ;i++){  
            cell = row10.createCell(i);  
            cell.setCellValue(data[6][i-1]);  
        }  
        
        Row row11 = sheet.createRow(11);   
        cell = row11.createCell(0);  
        cell.setCellValue("其中：保证金10%");  
        for(int i = 1;i <13 ;i++){  
            cell = row11.createCell(i);  
            cell.setCellValue(data[7][i-1]);  
        }  
        
        Row row12 = sheet.createRow(12);   
        
        Row row13 = sheet.createRow(13);   
        cell = row13.createCell(0);  
        cell.setCellValue("未签约");  
        for(int i = 1;i <13 ;i++){  
            cell = row13.createCell(i);  
            cell.setCellValue(data[8][i-1]);  
        }  
        
        Row row14 = sheet.createRow(14);   
        cell = row14.createCell(0);  
        cell.setCellValue("未签约面积");  
        for(int i = 1;i <13 ;i++){  
            cell = row14.createCell(i);  
            cell.setCellValue(data[9][i-1]);  
        }  
        
        Row row15 = sheet.createRow(15);   
        cell = row15.createCell(0);  
        cell.setCellValue("未签约金额");  
        for(int i = 1;i <13 ;i++){  
            cell = row15.createCell(i);  
            cell.setCellValue(data[10][i-1]);  
        }  
        
        Row row16 = sheet.createRow(16);   
        
        Row row17 = sheet.createRow(17);   
        cell = row17.createCell(0);  
        cell.setCellValue("已核对套数");  
        for(int i = 1;i <13 ;i++){  
            cell = row17.createCell(i);  
            cell.setCellValue(data[11][i-1]);  
        }  
        
        Row row18 = sheet.createRow(18);   
        cell = row18.createCell(0);  
        cell.setCellValue("已核对销售金额");  
        for(int i = 1;i <13 ;i++){  
            cell = row18.createCell(i);  
            cell.setCellValue(data[12][i-1]);  
        }  
        
        Row row19 = sheet.createRow(19);   
        cell = row19.createCell(0);  
        cell.setCellValue("已核对佣金");  
        for(int i = 1;i <13 ;i++){  
            cell = row19.createCell(i);  
            cell.setCellValue(data[13][i-1]);  
        }  
        
        Row row20 = sheet.createRow(20);   
        
        Row row21 = sheet.createRow(21);   
        cell = row21.createCell(0);  
        cell.setCellValue("未对账套数");  
        for(int i = 1;i <13 ;i++){  
            cell = row21.createCell(i);  
            cell.setCellValue(data[14][i-1]);  
        }  
        
        Row row22 = sheet.createRow(22);   
        cell = row22.createCell(0);  
        cell.setCellValue("未对账金额");  
        for(int i = 1;i <13 ;i++){  
            cell = row22.createCell(i);  
            cell.setCellValue(data[15][i-1]);  
        }  
        
        Row row23 = sheet.createRow(23);   
        cell = row23.createCell(0);  
        cell.setCellValue("未对账佣金");  
        for(int i = 1;i <13 ;i++){  
            cell = row23.createCell(i);  
            cell.setCellValue(data[16][i-1]);  
        }  
        
        //创建文件流   
        OutputStream stream = new FileOutputStream(excelPath);  
        //写入数据   
        wb.write(stream);  
        //关闭文件流   
        stream.close();  
    }  
}
