<%@ page language="java" import="java.util.*" pageEncoding="utf-8" import="servlet.Allfile"%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"> 
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Cache-Control" content="no-cache"> 
<meta http-equiv="Expires" content="0"> 
<title>表格下载</title> 
<link href="css/download.css" type="text/css" rel="stylesheet"> 
</head> 
<body> 
<div class="index">
    <div class="message">表格下载</div>
    <div id="darkbannerwrap"></div>
    <div class="file">
        <%
        ArrayList<String> files=Allfile.getFileName();%>
        <%for(int i=0;i<files.size();i++){ %>
        <form action="down" method="post">
		<input value=<%=files.get(i)%> style="width:100%;" type="submit" name="result">
		</form>
		<%} %>
	</div>
	 <hr class="hr15">
	<form action="down" method="post">
		<input value="下载全部文件" style="width:100%;color:#ffffff;background:#27A9E3;text-align:center" type="submit" name="result" >
	</form>
	
</div>
</body>
</html>