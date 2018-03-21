package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import ExcelManage.createTjbpm;
import ExcelManage.createXmzbfj;

public class Tjbpm {

	public static void makeTjbpm() throws IOException {
		List<String> visit=readData_visit.get("e:/test/原表.xls", 4);
		//System.out.println(list.size());
		/*for(int i=0;i<list.size();i++) {
			System.out.println(list.get(i));
		}*/
	    //获取当前月份
		List<String> list=readData.get("e:/test/原表.xls", 5);
		Calendar cal = Calendar.getInstance();
		int year = cal.get(Calendar.YEAR);
		int month = cal.get(Calendar.MONTH )+1;
		ArrayList<String> visit_data=new ArrayList<String>();
		for(int i=0;i<visit.size();i++) {
			int tempyear=Integer.valueOf(visit.get(i).split("<>")[0].split("-")[2]);
			int tempmonth=trans(visit.get(i).split("<>")[0].split("-")[1]);
			if(tempyear==year&&tempmonth==month) {
				visit_data.add(visit.get(i));
			}
		}
		
		ArrayList<String> data=new ArrayList<String>();
		for(int i=0;i<list.size();i++) {
			int tempyear=Integer.valueOf(list.get(i).split("<>")[1].split("-")[2]);
			int tempmonth=trans(list.get(i).split("<>")[1].split("-")[1]);
			if(tempyear==year&&tempmonth==month) {
				data.add(list.get(i));
			}
		}
		//全部来访记录  当月来访记录   全部成交记录   当月成交记录      
		makeTable(visit,visit_data,list,data,year,month);
        
	}
	
	public static void makeTable(List<String> visit,ArrayList<String> vpart,List<String> all,ArrayList<String> part,int year,int month) throws IOException {
		//获取全部销售员
		ArrayList<String> name=new ArrayList<String>();
		for(int i=0;i<all.size();i++) {
			String temp=all.get(i).split("<>")[43];
			if(temp.contains("（")) {
				//System.out.println(temp);
				String name1=temp.split("（")[0];
				//System.out.println(name1);
				String name2=temp.split("（")[1].split("）")[0];
				//System.out.println(name2);
				if(!name.contains(name1)) {
					name.add(name1);
				}
				if(!name.contains(name2)) {
					name.add(name2);
				}
			}else if(temp.contains("(")){
				//System.out.println(temp);
				String name1=temp.split("\\(")[0];
				//System.out.println(name1);
				String name2=temp.split("\\(")[1].split("\\)")[0];
				//System.out.println(name2);
				if(!name.contains(name1)) {
					name.add(name1);
				}
				if(!name.contains(name2)) {
					name.add(name2);
				}
			}else{
				if(!name.contains(temp)) {
					name.add(temp);
				}
			}
		}
		for(int i=name.size()-1;i>=0;i--) {
			if(name.get(i).equals("不计提")||name.get(i).equals("甲方")) {
				name.remove(i);
			}
		}
		for(int i=0;i<visit.size();i++) {
			String temp=visit.get(i).split("<>")[18];
			if(!name.contains(temp)) {
				name.add(temp);
			}
		}
		
		double[][] table=new double[name.size()+1][21];
		for(int i=0;i<name.size()+1;i++) {
			for(int j=0;j<21;j++) {
				table[i][j]=0;
			}
		}
		//本月的认购和签约
		for(int k=0;k<part.size();k++) {
			for(int j=0;j<name.size();j++) {
				if(part.get(k).split("<>")[43].equals(name.get(j))||(part.get(k).split("<>")[43].contains(name.get(j))&&part.get(k).split("<>")[43].contains("不计提"))||(part.get(k).split("<>")[43].contains(name.get(j))&&part.get(k).split("<>")[43].contains("甲方"))){
					table[j][2]+=1.0;
					table[j][3]+=Double.valueOf(part.get(k).split("<>")[11]);
					double temp=0;
					if(part.get(k).split("<>")[13].contains("+")) {
						temp=Double.valueOf(part.get(k).split("<>")[13].split("\\+")[0])+Double.valueOf(part.get(k).split("<>")[13].split("\\+")[1]);
						table[j][4]+=temp;
					}else {
						table[j][4]+=Double.valueOf(part.get(k).split("<>")[13]);
					}
					if(part.get(k).split("<>")[30].contains("-")) {
						table[j][11]+=1.0;
						table[j][12]+=Double.valueOf(part.get(k).split("<>")[11]);
						table[j][13]+=Double.valueOf(part.get(k).split("<>")[21]);
						
					}
				}else if(part.get(k).split("<>")[43].contains(name.get(j))) {
					table[j][2]+=0.5;
					table[j][3]+=0.5*Double.valueOf(part.get(k).split("<>")[11]);
					double temp=0;
					if(part.get(k).split("<>")[13].contains("+")) {
						temp=Double.valueOf(part.get(k).split("<>")[13].split("\\+")[0])+Double.valueOf(part.get(k).split("<>")[13].split("\\+")[1]);
						table[j][4]+=temp*0.5;
					}else {
						table[j][4]+=0.5*Double.valueOf(part.get(k).split("<>")[13]);
					}
					if(part.get(k).split("<>")[30].contains("-")) {
						table[j][11]+=0.5;
						table[j][12]+=0.5*Double.valueOf(part.get(k).split("<>")[11]);
						table[j][13]+=0.5*Double.valueOf(part.get(k).split("<>")[21]);
						
					}
				}
			}
		}

		//累计认购和签约
		for(int k=0;k<all.size();k++) {
			for(int j=0;j<name.size();j++) {
				if(all.get(k).split("<>")[43].equals(name.get(j))||(all.get(k).split("<>")[43].contains(name.get(j))&&all.get(k).split("<>")[43].contains("不计提"))||(all.get(k).split("<>")[43].contains(name.get(j))&&all.get(k).split("<>")[43].contains("甲方"))){
					table[j][6]+=1.0;
					table[j][7]+=Double.valueOf(all.get(k).split("<>")[11]);
					double temp=0;
					if(all.get(k).split("<>")[13].contains("+")) {
						temp=Double.valueOf(all.get(k).split("<>")[13].split("\\+")[0])+Double.valueOf(all.get(k).split("<>")[13].split("\\+")[1]);
						table[j][8]+=temp;
					}else {
						table[j][8]+=Double.valueOf(all.get(k).split("<>")[13]);
					}
					if(all.get(k).split("<>")[30].contains("-")) {
						table[j][16]+=1.0;
						table[j][17]+=Double.valueOf(all.get(k).split("<>")[11]);
						table[j][18]+=Double.valueOf(all.get(k).split("<>")[21]);
						
					}
				}else if(all.get(k).split("<>")[43].contains(name.get(j))) {
					table[j][6]+=0.5;
					table[j][7]+=0.5*Double.valueOf(all.get(k).split("<>")[11]);
					double temp=0;
					if(all.get(k).split("<>")[13].contains("+")) {
						temp=Double.valueOf(all.get(k).split("<>")[13].split("\\+")[0])+Double.valueOf(all.get(k).split("<>")[13].split("\\+")[1]);
						table[j][8]+=temp*0.5;
					}else {
						table[j][8]+=0.5*Double.valueOf(all.get(k).split("<>")[13]);
					}
					if(all.get(k).split("<>")[30].contains("-")) {
						table[j][16]+=0.5;
						table[j][17]+=0.5*Double.valueOf(all.get(k).split("<>")[11]);
						table[j][18]+=0.5*Double.valueOf(all.get(k).split("<>")[21]);
						
					}
				}
			}
		}
		
		//本月来访
		for(int k=0;k<vpart.size();k++) {
			for(int j=0;j<name.size();j++) {
				if(vpart.get(k).split("<>")[18].equals(name.get(j))){
							table[j][1]+=1.0;	
				}
			}
		}
		
		//累计来访
		for(int k=0;k<visit.size();k++) {
			for(int j=0;j<name.size();j++) {
				if(visit.get(k).split("<>")[18].equals(name.get(j))){
							table[j][0]+=1.0;	
				}
			}
		}
		
		//按当月认购套数排序
		for(int i=0;i<name.size();i++) {
			double[] temp=null;
			String str="";
			for(int j=0;j<name.size()-1-i;j++) {
				if(table[j][2]<table[j+1][2]) {
					temp=table[j];
					table[j]=table[j+1];
					table[j+1]=temp;	
					str=name.get(j);
					name.set(j, name.get(j+1));
					name.set(j+1, str);
				}
			}
		}
		
		//累计认购排名
		for(int i=0;i<name.size();i++) {
			double temp=1.0;
			for(int j=0;j<name.size();j++) {
				if(table[j][6]>table[i][6]) {
					temp+=1.0;
				}
			}
			table[i][10]=temp;
		}
		
		//当月签约排名
		for(int i=0;i<name.size();i++) {
			double temp=1.0;
			for(int j=0;j<name.size();j++) {
				if(table[j][11]>table[i][11]) {
					temp+=1.0;
				}
			}
			table[i][15]=temp;
		}
		
		//累计签约排名
		for(int i=0;i<name.size();i++) {
			double temp=1.0;
			for(int j=0;j<name.size();j++) {
				if(table[j][16]>table[i][16]) {
					temp+=1.0;
				}
			}
			table[i][20]=temp;
		}
		
		//合计
		for(int i=0;i<name.size();i++) {
			table[name.size()][0]+=table[i][0];
			table[name.size()][1]+=table[i][1];
			table[name.size()][2]+=table[i][2];
			table[name.size()][3]+=table[i][3];
			table[name.size()][4]+=table[i][4];
			table[name.size()][6]+=table[i][6];
			table[name.size()][7]+=table[i][7];
			table[name.size()][8]+=table[i][8];
			table[name.size()][11]+=table[i][11];
			table[name.size()][12]+=table[i][12];
			table[name.size()][13]+=table[i][13];
			table[name.size()][16]+=table[i][16];
			table[name.size()][17]+=table[i][17];
			table[name.size()][18]+=table[i][18];
		}
		
		//成交率和转签率
		for(int i=0;i<name.size()+1;i++) {
			if(table[i][1]!=0) {
				table[i][5]=table[i][2]/table[i][1];
			}
			if(table[i][0]!=0) {
				table[i][9]=table[i][6]/table[i][0];
			}
			if(table[i][2]!=0) {
				table[i][14]=table[i][11]/table[i][2];
			}
			if(table[i][6]!=0) {
				table[i][19]=table[i][16]/table[i][6];
			}
		}
		createTjbpm.create(name, table,year,month);
		
		/*for(int i=0;i<name.size()+1;i++) {
			//System.out.print(name.get(i)+"  ");
			System.out.print(table[i][0]+"  ");
			System.out.print(table[i][1]+"  ");
			System.out.print(table[i][2]+"  ");
			System.out.print(table[i][3]+"  ");
			System.out.print(table[i][4]+"  ");
			System.out.print(table[i][5]+"  ");
			System.out.print(table[i][6]+"  ");
			System.out.print(table[i][7]+"  ");
			System.out.print(table[i][8]+"  ");
			System.out.print(table[i][9]+"  ");
			System.out.print(table[i][10]+"  ");
			System.out.print(table[i][11]+"  ");
			System.out.print(table[i][12]+"  ");
			System.out.print(table[i][13]+"  ");
			System.out.print(table[i][14]+"  ");
			System.out.print(table[i][15]+"  ");
			System.out.print(table[i][16]+"  ");
			System.out.print(table[i][17]+"  ");
			System.out.print(table[i][18]+"  ");
			System.out.print(table[i][19]+"  ");
			System.out.println(table[i][20]);
		}*/
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
