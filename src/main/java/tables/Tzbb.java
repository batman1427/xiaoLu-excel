package tables;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

import ExcelManage.createTzbb;

public class Tzbb {
	public static void makeTzbb() throws IOException {
		List<String> list=readData.get("e:/test/原表.xls", 5);
		/*for(int i=0;i<list.size();i++) {
			System.out.println(list.get(i));
		}*/
		//找出全部年份
		int firstyear=Integer.valueOf(list.get(0).split("<>")[1].split("-")[2]);
		int lastyear=Integer.valueOf(list.get(list.size()-1).split("<>")[1].split("-")[2]);
        //按年份建表
		int start=0;
		int end=-1;
		for(int i=firstyear;i<=lastyear;i++) {
			//找出当年的所有数据
			for(int j=start;j<list.size();j++) {
				if(i==Integer.valueOf(list.get(j).split("<>")[1].split("-")[2])) {
					end++;
				}else{
					makeTable(list,start,end);
					start=end+1;
					j=list.size();
				}
			}
		}
		makeTable(list,start,end);
	}
	
	//按年份建表
	public static void makeTable(List<String> list,int start,int end) throws IOException {
		//按月份处理数据
		double[][] table=new double[17][12];
		for(int i=0;i<17;i++) {
			for(int j=0;j<12;j++) {
				table[i][j]=0;
			}
		}
		for(int i=start;i<=end;i++){
			int month=trans(list.get(i).split("<>")[1].split("-")[1])-1;
			//认购
			table[0][month]+=1.0;
			table[1][month]+=Double.valueOf(list.get(i).split("<>")[11]);
			double temp=0;
			if(list.get(i).split("<>")[13].contains("+")) {
				temp=Double.valueOf(list.get(i).split("<>")[13].split("\\+")[0])+Double.valueOf(list.get(i).split("<>")[13].split("\\+")[1]);
				table[2][month]+=temp;
			}else {
				table[2][month]+=Double.valueOf(list.get(i).split("<>")[13]);
			}
			//签约
			if(list.get(i).split("<>")[30].contains("-")) {
				table[3][month]+=1.0;
				table[4][month]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[5][month]+=Double.valueOf(list.get(i).split("<>")[21]);
				table[6][month]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
				
				double temp2=0;
				if(list.get(i).split("<>")[64].contains("*")) {
					temp2=Double.valueOf(list.get(i).split("<>")[64].split("\\*")[0].replace('%', '0'))/1000+Double.valueOf(list.get(i).split("<>")[64].split("\\*")[1]);
					table[7][month]+=temp2*Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
				}else {
					table[7][month]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60])*Double.valueOf(list.get(i).split("<>")[64]);
				}
			}else {
			//未签约
				table[8][month]+=1.0;
				table[9][month]+=Double.valueOf(list.get(i).split("<>")[11]);
				table[10][month]+=Double.valueOf(list.get(i).split("<>")[21]);
			}
			//已核对
			if(list.get(i).split("<>")[57].contains("-")) {
				table[11][month]+=1.0;
				table[12][month]+=Double.valueOf(list.get(i).split("<>")[21]);
				table[13][month]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
			}else {
			//未对账
				table[14][month]+=1.0;
				table[15][month]+=Double.valueOf(list.get(i).split("<>")[21]);
				table[16][month]+=Double.valueOf(list.get(i).split("<>")[21])*Double.valueOf(list.get(i).split("<>")[60]);
			}
		}
		/*for(int i=0;i<17;i++) {
			for(int j=0;j<12;j++) {
				System.out.print(table[i][j]+"  ");
			}
			System.out.println();
		}*/
		createTzbb.create(list.get(start).split("<>")[1].split("-")[2], table);
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
