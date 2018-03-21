package tables;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import servlet.Allfile;

public class controller {
    
	public static boolean maketable() throws IOException {
		ArrayList<String> files=Allfile.getFileName();
		if(files.size()==0) {
			return false;
		}else {
		// TODO Auto-generated method stub
        //台账报表     
		Tzbb.makeTzbb();
		//项目指标分解
		Xmzbfj.makeXmzbfj();
		//统计表排名
		Tjbpm.makeTjbpm();
		//项目销售回款情况
		Xmxshk.makeXmxshk();
		//统计表日报    项目蓄客情况 
		Tjbrb.makeTjbrb();
		//统计表日报的第二张表     项目销售情况
		Xsqk.makeXsqk();
		//统计表日报的第三张表     项目渠道成交对比
		Xmqd.makeXmqd();
		//统计表日报的第四张表    成交明细
		Cjmx.makeCjmx();
		}
		return true;
	}

	
}
