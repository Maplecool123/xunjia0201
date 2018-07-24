<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%response.charset="utf-8"%>
<!--#include file="UpLoadClass.asp"-->
<%
upload_url1=split(Request.ServerVariables("Path_Info"),"/edit/asp/aspupload.asp")
upload_url=upload_url1(0)&"/uploadfile/file/"
'upload_url="../Uploadfile/file/"
Server.ScriptTimeOut=10000
dim request2
set request2=New UpLoadClass
	request2.FileType="rar/zip/doc/docx/xls/xlsx/ppt/pptx/pdf/swf/fla/psd/txt/rm/rmvb/mpge/mpg/wmv/avi/mp3/flv"
	request2.maxsize=512000000
	request2.SavePath=upload_url
	request2.Open()	
	select case request2.form("imgFile"&"_Err")	
	case -1		
		Response.Write "<script language='javascript'>alert(""请选择上传文件!"");history.back();</script>"
	case 0	
		Response.write "<script type=text/javascript>parent.KE.plugin[""accessory""].insert("""&Request2.form("id")&""", """&trim(request2.savepath&request2.form("imgFile"))&""","""&Request2.form("imgTitle")&""");</script>"		
	case 1	
		Response.Write "<script language='javascript'>alert(""文件超过设定大小!"");history.back();</script>"	
	case 2	
		Response.Write "<script language='javascript'>alert(""文件类型不对!"");history.back();</script>"	
	case 3	
		Response.Write "<script language='javascript'>alert(""文件超过设定大小且文件类型不对!"");history.back();</script>"	
	case 4	
		Response.Write "<script language='javascript'>alert(""保存失败!"");history.back();</script>"		
	case else	
		Response.Write "<script language='javascript'>alert(""上传错误!"");history.back();</script>"	
	end select	
	set request2=nothing
%>