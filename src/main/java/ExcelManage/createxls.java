package ExcelManage;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createxls {

	public static void main(String[] args) throws IOException {  
        String path = "E:/";  
        String fileName = "测试";  
        String fileType = "xls";  
        ArrayList<String> list = read();
        writer(path, fileName, fileType,list);  
    }  
    
	public static ArrayList<String> read() throws IOException {
    	String pathname = "e:\\a.txt"; // 绝对路径或相对路径都可以，这里是绝对路径，写入文件时演示相对路径  
        File filename = new File(pathname); // 要读取以上路径的input。txt文件  
        InputStreamReader reader = new InputStreamReader(  
                new FileInputStream(filename)); // 建立一个输入流对象reader  
        BufferedReader br = new BufferedReader(reader); // 建立一个对象，它把文件内容转成计算机能读懂的语言 
        ArrayList<String>  s=new ArrayList<String>();
        String line = "";  
        while ((line = br.readLine())!= null) {   
        	s.add(line);
        }  
        return s;
    }
    public static void writer(String path, String fileName,String fileType,List<String> list) throws IOException {  
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
            sheet = (Sheet) wb.createSheet("sheet2");  
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
            sheet = (Sheet) wb.createSheet("sheet2");  
        }
        

        //循环写入行数据   
        for (int i = 0; i < 4; i++) {  
            Row row = (Row) sheet.createRow(i);  
            for(int j=0;j<50;j++) {
            //row.setHeight((short) 500); 
            row.createCell(j).setCellValue(Float.valueOf(list.get(j*4+i).split(" ")[7]));
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
