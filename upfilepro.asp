<!--#include file="Conn.asp" -->
<!--#include file="Const1.asp" -->
<!--#include file="upload.inc"-->
<link href='css/workspace.css' rel='stylesheet' type='text/css'>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<body bgcolor="#FFFFFF">
<%if request("chuan") = "yes" then
	dim upload,file,formName,formPath,iCount,filename,fileExt
	set upload=new upload_5xSoft ''建立上传对象
	iCount=0
	for each formName in upload.file ''列出所有上传了的文件
		set file=upload.file(formName)  ''生成一个文件对象
 		if file.filesize<10 then
 			response.write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			response.end
		 end if
 	
 		if file.filesize>uploadlimit*1024*10240 then
 			response.write "文件大小超过了限制 "&uploadlimit&" M　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			response.end
 		end if

 		fileExt=lcase(right(file.filename,4))

 		if fileEXT<>".xls" and fileEXT<>".xlsx" then
 			response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			response.end
 		end if 
 		randomize
 		ranNum=int(90000*rnd)+10000
 		filename=year(now)&month(now)&day(now)&ranNum&fileExt
		ok=filename
		filename="uploadfile/excel/"&filename
		if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
			file.SaveAs Server.mappath(FileName)   ''保存文件
			response.write "<script>parent.form1.ExName.value='"&ok&"'</script>"
			iCount=iCount+1
		end if
		set file=nothing
	next
	set upload=nothing  ''删除此对象
	response.write "<table width='100%' bgcolor='#ffffFF' border='0' cellspacing='1' cellpadding='2' style='margin-top:-8 px;'>"
	response.write "<tr><td height='29'>"
	response.write "<font size='+2' color=black>"
	response.write "Excel file uploads successful!"
	response.write "</font>"
	response.write "</td></tr></table>"
else%>
	<table border="0" cellspacing="1" cellpadding="4" bgcolor="#ffffFF" align="left" width="100%" style="margin-top:-8 px; margin-left:-10 px;">
	<form name="form" method="post" action="?chuan=yes" enctype="multipart/form-data" >
	<tr>
		<td height="10" width="100%">
		<input name="st" type="file" id="st" style="width:270px;" class="form_button">  
		<input name="Submit" type="submit" value="上传" class='form_submit' onSubmit="javascript:document.form.submit;">
		<input type="hidden" name="act" value="upload">
		</td>
	</tr>
	</form>
	</table>
<%end if%>
</body>