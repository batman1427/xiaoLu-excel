package servlet;

import java.io.File;
import java.util.ArrayList;

public class deletefiles {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
           delete();
	}

	 public static void delete() {
	      String path = "E:/test"; // 路径
	      File f = new File(path);
	     /* if (!f.exists()) {
	            System.out.println(path + " not exists");
	            return;
	       }*/

	       File fa[] = f.listFiles();
	       for (int i = 0; i < fa.length; i++) {
	           File fs = fa[i];
	           if (fs.isDirectory()) {
	              // System.out.println(fs.getName() + " [目录]");
	        	   fs.delete();
	           } else {
	              // System.out.println(fs.getName());
	        	   fs.delete();
	           }
	        }
	       System.out.println("文件全部删除");
	 }
}
