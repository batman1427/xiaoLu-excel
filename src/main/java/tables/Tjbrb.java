package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import ExcelManage.createTjbrb;
import ExcelManage.createTzbb;
import ExcelManage.createXmxshk;

public class Tjbrb {

	public static void makeTjbrb() throws IOException {
		List<String> zj=readData_zj.get("e:/test/原表.xls", 0);
		List<String> wt=readData_wt.get("e:/test/原表.xls", 2);
		List<String> cc=readData_cc.get("e:/test/原表.xls", 1);
		List<String> tel=readData_tel.get("e:/test/原表.xls", 3);
		List<String> visit=readData_visit.get("e:/test/原表.xls", 4);
		/*for(int i=0;i<wt.size();i++) {
			System.out.println(wt.get(i));
			//System.out.println(zj.get(i).split("<>").length);
		}*/
		Calendar cal = Calendar.getInstance();
		int year = cal.get(Calendar.YEAR);
		int month = cal.get(Calendar.MONTH )+1;
		int day=cal.get(Calendar.DATE);
		//System.out.println(day);
		makeTable1(year,month,day,zj,wt,cc,tel,visit);
	}
	
	//按年份建表
	public static void makeTable1(int year,int month,int day,List<String> zj,List<String> wt,List<String> cc,List<String> tel,List<String> visit) throws IOException {
		//找出所有中介类型
		ArrayList<String> zjtype=new ArrayList<String>();
		for(int i=0;i<zj.size();i++) {
			String temp=zj.get(i).split("<>")[2];
			if(!zjtype.contains(temp)) {
				zjtype.add(temp);
			}
		}
		
		double[][] table1=new double[zjtype.size()+10][10];
		for(int i=0;i<zjtype.size()+10;i++) {
			for(int j=0;j<10;j++) {
				table1[i][j]=0;
			}
		}
		
		//中介带访
		for(int i=0;i<zj.size();i++) {
			for(int j=0;j<zjtype.size();j++) {
				if(zj.get(i).split("<>")[2].equals(zjtype.get(j))) {
					//今日
					if(Integer.valueOf(zj.get(i).split("<>")[1].split("-")[2])==year&&trans(zj.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(zj.get(i).split("<>")[1].split("-")[0])==day) {
						table1[j][0]+=1.0;
						if(zj.get(i).split("<>").length>7&&zj.get(i).split("<>")[7].contains("-")) {
							table1[j][1]+=1.0;
						}
						if(zj.get(i).split("<>").length>9&&zj.get(i).split("<>")[9].equals("认购")) {
							table1[j][6]+=1.0;
						}
						if(zj.get(i).split("<>").length>10&&zj.get(i).split("<>")[10].contains("-")) {
							table1[j][8]+=1.0;
						}
					}
					//本月
					if(Integer.valueOf(zj.get(i).split("<>")[1].split("-")[2])==year&&trans(zj.get(i).split("<>")[1].split("-")[1])==month) {
						table1[j][2]+=1.0;
						if(zj.get(i).split("<>").length>7&&zj.get(i).split("<>")[7].contains("-")) {
							table1[j][3]+=1.0;
						}
					}
					//累计
					table1[j][4]+=1.0;
					if(zj.get(i).split("<>").length>7&&zj.get(i).split("<>")[7].contains("-")) {
						table1[j][5]+=1.0;
					}
					if(zj.get(i).split("<>").length>9&&zj.get(i).split("<>")[9].equals("认购")) {
						table1[j][7]+=1.0;
					}
					if(zj.get(i).split("<>").length>10&&zj.get(i).split("<>")[10].contains("-")) {
						table1[j][9]+=1.0;
					}
				}
			}
		}
		
		//外拓情况     拉访情况不明
		int wtstart=zjtype.size()+1;
		for(int i=0;i<wt.size();i++) {
			//今日
			if(Integer.valueOf(wt.get(i).split("<>")[1].split("-")[2])==year&&trans(wt.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(wt.get(i).split("<>")[1].split("-")[0])==day) {
				table1[wtstart][0]+=1.0;
				if(wt.get(i).split("<>").length>6&&wt.get(i).split("<>")[6].contains("-")) {
					table1[wtstart][1]+=1.0;
				}
				if(wt.get(i).split("<>").length>7&&wt.get(i).split("<>")[7].equals("认购")) {
					table1[wtstart][6]+=1.0;
				}
				if(wt.get(i).split("<>").length>8&&wt.get(i).split("<>")[8].contains("-")) {
					table1[wtstart][8]+=1.0;
				}
			}
			//本月
			if(Integer.valueOf(wt.get(i).split("<>")[1].split("-")[2])==year&&trans(wt.get(i).split("<>")[1].split("-")[1])==month) {
				table1[wtstart][2]+=1.0;
				if(wt.get(i).split("<>").length>6&&wt.get(i).split("<>")[6].contains("-")) {
					table1[wtstart][3]+=1.0;
				}
			}
			//累计
			table1[wtstart][4]+=1.0;
			if(wt.get(i).split("<>").length>6&&wt.get(i).split("<>")[6].contains("-")) {
				table1[wtstart][5]+=1.0;
			}
			if(wt.get(i).split("<>").length>7&&wt.get(i).split("<>")[7].equals("认购")) {
				table1[wtstart][7]+=1.0;
			}
			if(wt.get(i).split("<>").length>8&&wt.get(i).split("<>")[8].contains("-")) {
				table1[wtstart][9]+=1.0;
			}
		}
		
		//call客  均为call客户资源
		int ccstart=wtstart+3;
		for(int i=0;i<cc.size();i++) {
			//今日
			if(Integer.valueOf(cc.get(i).split("<>")[6].split("-")[2])==year&&trans(cc.get(i).split("<>")[6].split("-")[1])==month&&Integer.valueOf(cc.get(i).split("<>")[6].split("-")[0])==day) {
				table1[ccstart][0]+=1.0;
				if(cc.get(i).split("<>").length>11&&cc.get(i).split("<>")[11].contains("-")) {
					table1[ccstart][1]+=1.0;
				}
				if(cc.get(i).split("<>").length>13&&cc.get(i).split("<>")[13].equals("认购")) {
					table1[ccstart][6]+=1.0;
				}
				if(cc.get(i).split("<>").length>14&&cc.get(i).split("<>")[14].contains("-")) {
					table1[ccstart][8]+=1.0;
				}
			}
			//本月
			if(Integer.valueOf(cc.get(i).split("<>")[6].split("-")[2])==year&&trans(cc.get(i).split("<>")[6].split("-")[1])==month) {
				table1[ccstart][2]+=1.0;
				if(cc.get(i).split("<>").length>11&&cc.get(i).split("<>")[11].contains("-")) {
					table1[ccstart][3]+=1.0;
				}
			}
			//累计
			table1[ccstart][4]+=1.0;
			if(cc.get(i).split("<>").length>11&&cc.get(i).split("<>")[11].contains("-")) {
				table1[ccstart][5]+=1.0;
			}
			if(cc.get(i).split("<>").length>13&&cc.get(i).split("<>")[13].equals("认购")) {
				table1[ccstart][7]+=1.0;
			}
			if(cc.get(i).split("<>").length>14&&cc.get(i).split("<>")[14].contains("-")) {
				table1[ccstart][9]+=1.0;
			}
		}
		
		//来电情况
		int telstart=ccstart+2;
		for(int i=0;i<tel.size();i++) {
			//今日
			if(Integer.valueOf(tel.get(i).split("<>")[1].split("-")[2])==year&&trans(tel.get(i).split("<>")[1].split("-")[1])==month&&Integer.valueOf(tel.get(i).split("<>")[1].split("-")[0])==day) {
				table1[telstart][0]+=1.0;
				if(tel.get(i).split("<>").length>12&&tel.get(i).split("<>")[12].contains("-")) {
					table1[telstart][1]+=1.0;
				}
				if(tel.get(i).split("<>").length>13&&tel.get(i).split("<>")[13].equals("认购")) {
					table1[telstart][6]+=1.0;
				}
				if(tel.get(i).split("<>").length>14&&tel.get(i).split("<>")[14].contains("-")) {
					table1[telstart][8]+=1.0;
				}
			}
			//本月
			if(Integer.valueOf(tel.get(i).split("<>")[1].split("-")[2])==year&&trans(tel.get(i).split("<>")[1].split("-")[1])==month) {
				table1[telstart][2]+=1.0;
				if(tel.get(i).split("<>").length>12&&tel.get(i).split("<>")[12].contains("-")) {
					table1[telstart][3]+=1.0;
				}
			}
			//累计
			table1[telstart][4]+=1.0;
			if(tel.get(i).split("<>").length>12&&tel.get(i).split("<>")[12].contains("-")) {
				table1[telstart][5]+=1.0;
			}
			if(tel.get(i).split("<>").length>13&&tel.get(i).split("<>")[13].equals("认购")) {
				table1[telstart][7]+=1.0;
			}
			if(tel.get(i).split("<>").length>14&&tel.get(i).split("<>")[14].contains("-")) {
				table1[telstart][9]+=1.0;
			}
		}
		
		//来访情况  新客户来访
		int visitstart=telstart+1;
		for(int i=0;i<visit.size();i++) {
		   if(Double.valueOf(visit.get(i).split("<>")[3])==1.0){
			//今日
			if(Integer.valueOf(visit.get(i).split("<>")[0].split("-")[2])==year&&trans(visit.get(i).split("<>")[0].split("-")[1])==month&&Integer.valueOf(visit.get(i).split("<>")[0].split("-")[0])==day) {
				table1[visitstart][0]+=1.0;
				if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
					table1[visitstart][1]+=1.0;
				}
				if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
					table1[visitstart][6]+=1.0;
				}
				if(visit.get(i).split("<>").length>19&&visit.get(i).split("<>")[19].contains("-")) {
					table1[visitstart][8]+=1.0;
				}
			}
			//本月
			if(Integer.valueOf(visit.get(i).split("<>")[0].split("-")[2])==year&&trans(visit.get(i).split("<>")[0].split("-")[1])==month) {
				table1[visitstart][2]+=1.0;
				if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
					table1[visitstart][3]+=1.0;
				}
			}
			//累计
			table1[visitstart][4]+=1.0;
			if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
				table1[visitstart][5]+=1.0;
			}
			if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
				table1[visitstart][7]+=1.0;
			}
			if(visit.get(i).split("<>").length>19&&visit.get(i).split("<>")[19].contains("-")) {
				table1[visitstart][9]+=1.0;
			}
		   }
		}
		
		// 老客户来访
		int oldstart=visitstart+1;
		for(int i=0;i<visit.size();i++) {
		   if(Double.valueOf(visit.get(i).split("<>")[3])>1.0){
			//今日
			if(Integer.valueOf(visit.get(i).split("<>")[0].split("-")[2])==year&&trans(visit.get(i).split("<>")[0].split("-")[1])==month&&Integer.valueOf(visit.get(i).split("<>")[0].split("-")[0])==day) {
				table1[oldstart][0]+=1.0;
				if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
					table1[oldstart][1]+=1.0;
				}
				if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
					table1[oldstart][6]+=1.0;
				}
				if(visit.get(i).split("<>").length>19&&visit.get(i).split("<>")[19].contains("-")) {
					table1[oldstart][8]+=1.0;
				}
			}
			//本月
			if(Integer.valueOf(visit.get(i).split("<>")[0].split("-")[2])==year&&trans(visit.get(i).split("<>")[0].split("-")[1])==month) {
				table1[oldstart][2]+=1.0;
				if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
					table1[oldstart][3]+=1.0;
				}
			}
			//累计
			table1[oldstart][4]+=1.0;
			if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
				table1[oldstart][5]+=1.0;
			}
			if(visit.get(i).split("<>").length>16&&visit.get(i).split("<>")[16].equals("认购")) {
				table1[oldstart][7]+=1.0;
			}
			if(visit.get(i).split("<>").length>19&&visit.get(i).split("<>")[19].contains("-")) {
				table1[oldstart][9]+=1.0;
			}
		   }
		}
		
		//小计
		for(int i=0;i<10;i++) {
			table1[telstart-1][i]=table1[telstart-2][i]+table1[telstart-3][i];
		}
		
		//合计
		for(int i=0;i<10;i++) {
			for(int j=0;j<oldstart+1;j++) {
				//去掉小计
				if(j!=oldstart-3) {
				table1[oldstart+1][i]+=table1[j][i];
				}
			}
		}
		/*for(int i=0;i<zjtype.size()+10;i++) {
			System.out.print(table1[i][0]+"  ");
			System.out.print(table1[i][1]+"  ");
			System.out.print(table1[i][2]+"  ");
			System.out.print(table1[i][3]+"  ");
			System.out.print(table1[i][4]+"  ");
			System.out.print(table1[i][5]+"  ");
			System.out.print(table1[i][6]+"  ");
			System.out.print(table1[i][7]+"  ");
			System.out.print(table1[i][8]+"  ");
			System.out.println(table1[i][9]+"  ");
		}*/
		createTjbrb.create(table1,year,month,day,zjtype.size(), zjtype);
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
