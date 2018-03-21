<%@ page language="java" import="java.util.*" pageEncoding="utf-8"%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"> 
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Cache-Control" content="no-cache"> 
<meta http-equiv="Expires" content="0"> 
<title>表格汇总</title> 
<link href="css/index.css" type="text/css" rel="stylesheet"> 
</head> 
<body> 
<div class="index">
    <div class="message">表格汇总</div>
    <div id="darkbannerwrap"></div>
    
    <form action="index" method="post" enctype="multipart/form-data">
        <input type="text" value="基础表格:" readonly="readonly" style="border:none;font-weight:bold;">
        <hr class="hr10">
        <input type="file" name="file" size="50"  />
        <hr class="hr15">
		<input value="提交" style="width:100%;" type="submit">
		<hr class="hr20">
	</form>
</div>
</body>
</html>