package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import ExcelManage.createTjbrb;
import ExcelManage.createXmqd;
import ExcelManage.createXsqk;

public class Xmqd {

	public static void makeXmqd() throws IOException {
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
	
	public static void makeTable(int year,int month,int day,List<String> list) throws IOException {
		
		double[][] table=new double[7][16];
		for(int i=0;i<7;i++) {
			for(int j=0;j<16;j++) {
				table[i][j]=0;
			}
		}
		for(int i=0;i<list.size();i++) {
			if(list.get(i).split("<>")[52].contains("自然来访")) {
				table[0][2]+=1.0;
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[0][3]+=temp;
				}else {
					table[0][3]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
				if(list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
					table[0][0]+=1.0;
					if(list.get(i).split("<>")[13].contains("+")) {
						temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
						table[0][1]+=temp;
					}else {
						table[0][1]+=Double.valueOf(list.get(i).split("<>")[13]);
					}
					
				}
			}else if(list.get(i).split("<>")[52].contains("中原")) {
				table[1][2]+=1.0;
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[1][3]+=temp;
				}else {
					table[1][3]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
				if(list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
					table[1][0]+=1.0;
					if(list.get(i).split("<>")[13].contains("+")) {
						temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
						table[1][1]+=temp;
					}else {
						table[1][1]+=Double.valueOf(list.get(i).split("<>")[13]);
					}
					
				}
				
			}else if(list.get(i).split("<>")[52].contains("小鹿")) {
				table[2][2]+=1.0;
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[2][3]+=temp;
				}else {
					table[2][3]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
				if(list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
					table[2][0]+=1.0;
					if(list.get(i).split("<>")[13].contains("+")) {
						temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
						table[2][1]+=temp;
					}else {
						table[2][1]+=Double.valueOf(list.get(i).split("<>")[13]);
					}
					
				}
				
			}else if(list.get(i).split("<>")[52].contains("老带新")) {
				table[3][2]+=1.0;
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[3][3]+=temp;
				}else {
					table[3][3]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
				if(list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
					table[3][0]+=1.0;
					if(list.get(i).split("<>")[13].contains("+")) {
						temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
						table[3][1]+=temp;
					}else {
						table[3][1]+=Double.valueOf(list.get(i).split("<>")[13]);
					}
					
				}
				
			}else if(list.get(i).split("<>")[52].contains("全民营销")) {
				table[4][2]+=1.0;
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[4][3]+=temp;
				}else {
					table[4][3]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
				if(list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
					table[4][0]+=1.0;
					if(list.get(i).split("<>")[13].contains("+")) {
						temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
						table[4][1]+=temp;
					}else {
						table[4][1]+=Double.valueOf(list.get(i).split("<>")[13]);
					}
					
				}
				
			}else{
				table[5][2]+=1.0;
				double temp=0;
				if(list.get(i).split("<>")[13].contains("+")) {
					temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
					table[5][3]+=temp;
				}else {
					table[5][3]+=Double.valueOf(list.get(i).split("<>")[13]);
				}
				if(list.get(i).split("<>")[1].contains("-")&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[2])==year&&trans(list.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(list.get(i).split("<>")[1].split("-")[0])==day) {
					table[5][0]+=1.0;
					if(list.get(i).split("<>")[13].contains("+")) {
						temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
						table[5][1]+=temp;
					}else {
						table[5][1]+=Double.valueOf(list.get(i).split("<>")[13]);
					}
					
				}
				
			}
		}
		for(int i=0;i<12;i++) {
			for(int j=0;j<6;j++) {
				table[6][i]+=table[j][i];
			}
		}
		
		/*for(int i=0;i<7;i++) {
			for(int j=0;j<11;j++) {
				System.out.print(table[i][j]+"  ");
			}
			System.out.println(table[i][11]);
		}*/
		createXmqd.create(table,year,month,day);
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
