package tables;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readData {
	public static List<String> get(String filename,int num) throws IOException {
		List<String> list=readExcel(filename,num);  
		return list;
		
	}

	public static List<String>  readExcel(String filename,int page) throws IOException {
		
        //获得Workbook工作薄对象  
        Workbook workbook = getWorkBook(filename);  
        //创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回  
        List<String> list = new ArrayList<String>();  
        if(workbook != null){  
            int sheetNum = page; 
                //获得当前sheet工作表  
                Sheet sheet = (Sheet) workbook.getSheetAt(sheetNum);  
                if(sheet == null){  
                     
                }  
                //获得当前sheet的开始行  
                int firstRowNum  = sheet.getFirstRowNum();  
                //获得当前sheet的结束行  
                int lastRowNum = sheet.getLastRowNum();  
                //循环除了第一行的所有行  
                for(int rowNum = firstRowNum+1;rowNum <= lastRowNum;rowNum++){  
                    //获得当前行  
                	  Row row = sheet.getRow(rowNum);  
                      Cell temp = row.getCell(1); 
                      if(row == null||!String.valueOf(temp).contains("-")){  
                          continue;  
                      }  
                    //获得当前行的开始列  
                    int firstCellNum = 0;  
                    //获得当前行的列数  
                    int lastCellNum = 75;  
                    String result="";  
                    //循环当前行  
                    for(int cellNum = firstCellNum; cellNum < lastCellNum;cellNum++){  
                    	Cell cell = row.getCell(cellNum); 
                    	if (cell.getCellType() == Cell.CELL_TYPE_FORMULA && cell!=null) {
                    		try {
                    		    result += String.valueOf(cell.getNumericCellValue())+"<>";}
                    		catch(Exception e){
                    			result += String.valueOf(cell)+"<>"; 
                    			continue;
                    		}
                         }else {
                
                           result += String.valueOf(cell)+"<>";  
                        }
                    }  
                    list.add(result);  
                }  
             
            workbook.close();  
        }
		return list;  
	}
	
	//获取表格内容xls和xlsx
	public static Workbook getWorkBook(String filename) {  
        //创建Workbook工作薄对象，表示整个excel  
        Workbook workbook = null;  
        try {  
            //获取excel文件的io流 
        	InputStream instream = new FileInputStream(filename);  
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if(filename.endsWith("xls")){  
                //2003  
                workbook = new HSSFWorkbook(instream);  
            }else if(filename.endsWith("xlsx")){  
                //2007  
                workbook = new XSSFWorkbook(instream);  
            }  
        } catch (IOException e) {  
        }  
        return workbook;  
    }  
}
