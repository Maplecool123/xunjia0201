<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%response.charset="utf-8"%>
<!--#include file="UpLoadClass.asp"-->
<!--#include file="JSON_2.0.4.asp"-->
<%
upload_url1=split(Request.ServerVariables("Path_Info"),"/edit/asp/upload_json.asp")

upload_url=upload_url1(0)&"/Uploadfile/pic/"
'upload_url="../Uploadfile/pic/"
Server.ScriptTimeOut=5000
dim request2
set request2=New UpLoadClass
	request2.FileType="jpg/jpge/gif/bmp/png"
	request2.maxsize=10240000
	request2.SavePath=upload_url
	request2.Open()
	select case request2.form("imgFile"&"_Err")	
	case -1	
		call showError("请选择上传文件!")
	case 0
		Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
		Set hash = jsObject()
		hash("error") = 0
		hash("url") = trim(request2.savepath&request2.form("imgFile"))
		hash.Flush
		Response.End
	case 1	
		call showError("文件超过设定大小!")
	case 2	
		call showError("文件类型不对!")
	case 3
		call showError("文件超过设定大小且文件类型不对!")
	case 4	
		call showError("保存失败!")
	case else
		call showError("上传错误!")
	end select
	set request2=nothing

Function showError(message)
	Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
	Dim hash
	Set hash = jsObject()
	hash("error") = 1
	hash("message") = message
	hash.Flush
	Response.End
End Function
%>