package servlet;

import java.io.File;
import java.util.ArrayList;

public class Allfile {

	 public static void main(String[] args) {
	        getFileName();
	    }

	 public static ArrayList<String> getFileName() {
	      String path = "E:/test"; // 路径
	      File f = new File(path);
	     /* if (!f.exists()) {
	            System.out.println(path + " not exists");
	            return;
	       }*/

	       File fa[] = f.listFiles();
	       ArrayList<String>  files=new ArrayList<String>();
	       for (int i = 0; i < fa.length; i++) {
	           File fs = fa[i];
	           if (fs.isDirectory()) {
	              // System.out.println(fs.getName() + " [目录]");
	           } else {
	              // System.out.println(fs.getName());
	        	   files.add(fs.getName());
	           }
	        }
	       files.remove("Alltables.zip");
	       return files;
	 }
}
