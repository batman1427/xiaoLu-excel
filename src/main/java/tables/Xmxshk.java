package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import ExcelManage.createTzbb;
import ExcelManage.createXmxshk;

public class Xmxshk {

	public static void makeXmxshk() throws IOException {
		List<String> list=readData.get("e:/test/原表.xls", 5);
		//找出全部年份
		int firstyear=Integer.valueOf(list.get(0).split("<>")[1].split("-")[2]);
		int lastyear=Integer.valueOf(list.get(list.size()-1).split("<>")[1].split("-")[2]);
        //按年份建表
		makeTable(list,firstyear,lastyear);
	}
	
	//按年份建表
	public static void makeTable(List<String> list,int start,int end) throws IOException {
		int years=end-start+1;
		//按年处理数据
		double[][] table=new double[(years+1)*3+1][23];
		for(int i=0;i<(years+1)*3+1;i++) {
			for(int j=0;j<23;j++) {
				table[i][j]=0;
			}
		}
		
		//根据年份和住房类型划分
		ArrayList<String> zhuzhai=new ArrayList<String>();
		zhuzhai.add("");
		ArrayList<String> shangpu=new ArrayList<String>();
		shangpu.add("商铺");
		ArrayList<String> gongyu=new ArrayList<String>();
		gongyu.add("公寓");
		for(int i=0;i<list.size();i++) {
			//获取年份
			int  year=Integer.valueOf(list.get(i).split("<>")[1].split("-")[2]);
			int temp=0;
			String type=list.get(i).split("<>")[3];
			if(zhuzhai.contains(type)) {
				temp=year-start;
			}else if(shangpu.contains(type)) {
				temp=year-start+years+1;
			}else if(gongyu.contains(type)) {
				temp=year-start+years*2+2;
			}
			//认购
			table[temp][0]+=1.0;
			table[temp][1]+=Double.valueOf(list.get(i).split("<>")[11]);
			double aa=0;
			if(list.get(i).split("<>")[13].contains("+")) {
				aa=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
				table[temp][2]+=aa;
			}else {
				table[temp][2]+=Double.valueOf(list.get(i).split("<>")[13]);
			}
			//签约
			if(list.get(i).split("<>")[30].contains("-")) {
				table[temp][3]+=1.0;
				table[temp][4]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[temp][5]+=Double.valueOf(list.get(i).split("<>")[21]);
				
			}else {
			//未签约
				table[temp][6]+=1.0;
				table[temp][7]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[temp][8]+=Double.valueOf(list.get(i).split("<>")[13]);
			}
			//三部分累加等于总签约
			if(list.get(i).split("<>")[30].contains("-")) {
				//已下款已提交报告   清款日期+佣金结算单
				if(list.get(i).split("<>")[30].contains("-")&&list.get(i).split("<>")[57].contains("-")) {
					table[temp][10]+=1.0;
					table[temp][11]+=Double.valueOf(list.get(i).split("<>")[11]);
					table[temp][12]+=Double.valueOf(list.get(i).split("<>")[21]);
					table[temp][13]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
				//已下款未提交报告    清款未结佣
				}else if(list.get(i).split("<>")[30].contains("-")&&!list.get(i).split("<>")[57].contains("-")) {
					table[temp][15]+=1.0;
					table[temp][16]+=Double.valueOf(list.get(i).split("<>")[11]);
					table[temp][17]+=Double.valueOf(list.get(i).split("<>")[21]);
					table[temp][18]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
				//未下款
				}else {
					table[temp][19]+=1.0;
					table[temp][20]+=Double.valueOf(list.get(i).split("<>")[11]);
					table[temp][21]+=Double.valueOf(list.get(i).split("<>")[21]);
					table[temp][22]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
				}
			}
			
			
		}
		
		//住宅合计
		for(int i=0;i<23;i++) {
			for(int j=0;j<years;j++) {
				table[years][i]+=table[j][i];
			}
		}
		//商铺合计
		for(int i=0;i<23;i++) {
			for(int j=years+1;j<years*2+1;j++) {
				table[years*2+1][i]+=table[j][i];
			}
		}
		//公寓合计
		for(int i=0;i<23;i++) {
			for(int j=years*2+2;j<years*3+2;j++) {
				table[years*3+2][i]+=table[j][i];
			}
		}
		//总合计
		for(int i=0;i<23;i++) {
			for(int j=0;j<3;j++) {
				table[years*3+3][i]+=table[j*(years+1)+3][i];
			}
		}
		/*
		for(int i=0;i<(years+1)*3+1;i++) {
			for(int j=0;j<23;j++) {
				System.out.print(table[i][j]+"  ");
			}
			System.out.println();
		}*/
		createXmxshk.create(table,years,start);
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
