package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import ExcelManage.createTzbb;
import ExcelManage.createXmzbfj;

public class Xmzbfj {

	public static void makeXmzbfj() throws IOException {
		// TODO Auto-generated method stub
		List<String> list=readData.get("e:/test/原表.xls", 5);
		/*for(int i=0;i<list.size();i++) {
			System.out.println(list.get(i));
		}*/
	    //获取当前月份
		Calendar cal = Calendar.getInstance();
		int year = cal.get(Calendar.YEAR);
		int month = cal.get(Calendar.MONTH )+1;
		ArrayList<String> data=new ArrayList<String>();
		for(int i=0;i<list.size();i++) {
			int tempyear=Integer.valueOf(list.get(i).split("<>")[1].split("-")[2]);
			int tempmonth=trans(list.get(i).split("<>")[1].split("-")[1]);
			if(tempyear==year&&tempmonth==month) {
				data.add(list.get(i));
			}
		}
		makeTable(data,year,month);
	}
	
	public static void makeTable(ArrayList<String> data,int year,int month) throws IOException {
		//获取该月所有销售员
		ArrayList<String> name=new ArrayList<String>();
		for(int i=0;i<data.size();i++) {
			String temp=data.get(i).split("<>")[43];
			if(!name.contains(temp)) {
				name.add(temp);
			}
		}
		double[][] table=new double[name.size()][5];
		for(int i=0;i<name.size();i++) {
			for(int j=0;j<5;j++) {
				table[i][j]=0;
			}
		}
		for(int k=0;k<data.size();k++) {
			for(int j=0;j<name.size();j++) {
				if(data.get(k).split("<>")[43].equals(name.get(j))){
					table[j][0]+=1.0;
					double temp=0;
					if(data.get(k).split("<>")[13].contains("+")) {
						temp=Double.valueOf(data.get(k).split("<>")[13].split("\\+")[0])+Double.valueOf(data.get(k).split("<>")[13].split("\\+")[1]);
						table[j][1]+=temp;
					}else {
						table[j][1]+=Double.valueOf(data.get(k).split("<>")[13]);
					}
					if(data.get(k).split("<>")[30].contains("-")) {
						table[j][2]+=1.0;
						table[j][3]+=Double.valueOf(data.get(k).split("<>")[21]);
						table[j][4]+=Double.valueOf(data.get(k).split("<>")[21])*Double.valueOf(data.get(k).split("<>")[60]);
					}
				}
			}
		}
		/*for(int i=0;i<name.size();i++) {
			System.out.print(name.get(i)+"  ");
			System.out.print(table[i][0]+"  ");
			System.out.print(table[i][1]+"  ");
			System.out.print(table[i][2]+"  ");
			System.out.print(table[i][3]+"  ");
			System.out.println(table[i][4]);
		}*/
		createXmzbfj.create(name, table,year,month);
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
