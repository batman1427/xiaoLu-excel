package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import ExcelManage.createCjmx;
import ExcelManage.createTjbrb;
import ExcelManage.createXsqk;

public class Cjmx {

	public static void makeCjmx() throws IOException {
		List<String> list=readData.get("e:/test/原表.xls", 5);
		/*for(int i=0;i<list.size();i++) {
			System.out.println(list.get(i));
			//System.out.println(zj.get(i).split("<>").length);
		}*/		
		makeTable(list);
	}
	
	//按年份建表
	public static void makeTable(List<String> list) throws IOException {
		//找出全部的符合条件的成交记录
        ArrayList<String> result=new  ArrayList<String>();
        for(int i=0;i<list.size();i++) {
        	if(list.get(i).split("<>")[30].contains("-")) {
        		if(list.get(i).split("<>")[52].contains("老带新")||list.get(i).split("<>")[52].contains("全民营销")) {
        			result.add(list.get(i));
        		}
        	}
        }	
        //System.out.println(result.size());
		String[][] table=new String[result.size()][9];
		for(int i=0;i<result.size();i++) {
			for(int j=0;j<9;j++) {
				table[i][j]="";
			}
		}
		
		for(int i=0;i<result.size();i++) {
			table[i][0]=String.valueOf(result.get(i).split("<>")[30].split("-")[2]+"/"+trans(result.get(i).split("<>")[30].split("-")[1])+"/"+result.get(i).split("<>")[30].split("-")[0]);
			table[i][1]=String.valueOf(result.get(i).split("<>")[53]);
			table[i][2]=String.valueOf(result.get(i).split("<>")[54]);
			table[i][3]=String.valueOf(result.get(i).split("<>")[8]);
			table[i][4]=String.valueOf(result.get(i).split("<>")[7]);
			table[i][5]=String.valueOf(result.get(i).split("<>")[21]);
			table[i][6]=String.valueOf(result.get(i).split("<>")[43]);
			//table[i][7]=String.valueOf(result.get(i).split("<>")[1]);
			table[i][8]=String.valueOf(result.get(i).split("<>")[52]);
		}
		
		/*for(int i=0;i<result.size();i++) {
			for(int j=0;j<8;j++) {
				System.out.print(table[i][j]+"  ");
			}
			System.out.println(table[i][8]);
		}*/
		createCjmx.create(table);
	}
	
	//月份换算
	public static int trans(String month) {
		int result=0;
		switch(month) {
			case "一月":result=1;
			break;
			case "二月":result=2;
			break;
			case "三月":result=3;
			break;
			case "四月":result=4;
			break;
			case "五月":result=5;
			break;
			case "六月":result=6;
			break;
			case "七月":result=7;
			break;
			case "八月":result=8;
			break;
			case "九月":result=9;
			break;
			case "十月":result=10;
			break;
			case "十一月":result=11;
			break;
			case "十二月":result=12;
		}
		return result;
	}

}
