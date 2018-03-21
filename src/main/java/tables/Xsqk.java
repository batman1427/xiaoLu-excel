package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import ExcelManage.createTjbrb;
import ExcelManage.createXsqk;

public class Xsqk {

	public static void makeXsqk() throws IOException {
		List<String> list=readData.get("e:/test/原表.xls", 5);
		/*for(int i=0;i<list.size();i++) {
			System.out.println(list.get(i));
			//System.out.println(zj.get(i).split("<>").length);
		}*/
		Calendar cal = Calendar.getInstance();
		int year = cal.get(Calendar.YEAR);
		int month = cal.get(Calendar.MONTH )+1;
		int day=cal.get(Calendar.DATE);
		//System.out.println(day);
		makeTable(year,month,day,list);
	}
	
	//按年份建表
	public static void makeTable(int year,int month,int day,List<String> list) throws IOException {
		ArrayList<String> zhuzhai=new ArrayList<String>();
		zhuzhai.add("");
		ArrayList<String> shangpu=new ArrayList<String>();
		shangpu.add("商铺");
		
		double[][] table=new double[20][12];
		for(int i=0;i<20;i++) {
			for(int j=0;j<12;j++) {
				table[i][j]=0;
			}
		}
		for(int i=0;i<list.size();i++) {
			String type=list.get(i).split("<>")[3];
			//认筹
			if(list.get(i).split("<>")[2].contains("-")&&Integer.valueOf(list.get(i).split("<>")[2].split("-")[2])==year&&trans(list.get(i).split("<>")[2].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[2].split("-")[0])==day) {
				table[0][0]+=1.0;
				table[0][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[0][2]+=temp;
				}else {
					table[0][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(list.get(i).split("<>")[2].contains("-")&&Integer.valueOf(list.get(i).split("<>")[2].split("-")[2])==year&&trans(list.get(i).split("<>")[2].split("-")[1])==month) {
				table[1][0]+=1.0;
				table[1][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[1][2]+=temp;
				}else {
					table[1][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(list.get(i).split("<>")[2].contains("-")) {
				table[2][0]+=1.0;
				table[2][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[2][2]+=temp;
				}else {
					table[2][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			
			//住宅认购
			if(zhuzhai.contains(type)&&list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
				table[3][0]+=1.0;
				table[3][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[3][2]+=temp;
				}else {
					table[3][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(zhuzhai.contains(type)&&list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month) {
				table[4][0]+=1.0;
				table[4][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[4][2]+=temp;
				}else {
					table[4][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(zhuzhai.contains(type)&&list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year) {
				table[5][0]+=1.0;
				table[5][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[5][2]+=temp;
				}else {
					table[5][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(zhuzhai.contains(type)&&list.get(i).split("<>")[1].contains("-")) {
				table[6][0]+=1.0;
				table[6][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[6][2]+=temp;
				}else {
					table[6][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			
			//商业认购
			if(shangpu.contains(type)&&list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
				table[7][0]+=1.0;
				table[7][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[7][2]+=temp;
				}else {
					table[7][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(shangpu.contains(type)&&list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month) {
				table[8][0]+=1.0;
				table[8][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[8][2]+=temp;
				}else {
					table[8][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(shangpu.contains(type)&&list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year) {
				table[9][0]+=1.0;
				table[9][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[9][2]+=temp;
				}else {
					table[9][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			if(shangpu.contains(type)&&list.get(i).split("<>")[1].contains("-")) {
				table[10][0]+=1.0;
				table[10][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[10][2]+=temp;
				}else {
					table[10][2]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
			}
			
			//签约
			if(list.get(i).split("<>")[30].contains("-")&&Integer.valueOf(list.get(i).split("<>")[30].split("-")[2])==year&&trans(list.get(i).split("<>")[30].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[30].split("-")[0])==day) {
				table[11][0]+=1.0;
				table[11][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[11][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			if(list.get(i).split("<>")[30].contains("-")&&Integer.valueOf(list.get(i).split("<>")[30].split("-")[2])==year&&trans(list.get(i).split("<>")[30].split("-")[1])==month) {
				table[12][0]+=1.0;
				table[12][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[12][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			if(list.get(i).split("<>")[30].contains("-")&&Integer.valueOf(list.get(i).split("<>")[30].split("-")[2])==year) {
				table[13][0]+=1.0;
				table[13][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[13][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			if(list.get(i).split("<>")[30].contains("-")) {
				table[14][0]+=1.0;
				table[14][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[14][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			
			//未签约
			if(!list.get(i).split("<>")[30].contains("-")) {
				table[15][0]+=1.0;
				table[15][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[15][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			
			//已签约已下款
			if(list.get(i).split("<>")[30].contains("-")&&list.get(i).split("<>")[38].contains("-")) {
				table[16][0]+=1.0;
				table[16][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[16][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			
			//已下款已结佣
			if(list.get(i).split("<>")[38].contains("-")&&list.get(i).split("<>")[57].contains("-")) {
				table[17][0]+=1.0;
				table[17][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[17][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			
			//已签约未下款
			if(list.get(i).split("<>")[30].contains("-")&&!list.get(i).split("<>")[38].contains("-")) {
				table[18][0]+=1.0;
				table[18][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[18][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			
			//已签约已下款
			if(list.get(i).split("<>")[38].contains("-")&&!list.get(i).split("<>")[57].contains("-")) {
				table[19][0]+=1.0;
				table[19][1]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[19][2]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
		}
		
		
		
	/*	for(int i=0;i<20;i++) {
			for(int j=0;j<2;j++) {
				System.out.print(table[i][j]+"  ");
			}
			System.out.println(table[i][2]);
		}*/
		createXsqk.create(table,year,month,day);
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
