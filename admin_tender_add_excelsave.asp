<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-99=0 then
call WhereTable("gnome-ccdesktop.png","正在导入数据")

djh = request.Form("djh")
ExName = Request.Form("ExName")
ExTName = Request.Form("ExTName")
ifupdate = Request.Form("ifupdate")
Dim conn2
'On Error Resume Next
Server.ScriptTimeOut = 999999
set conn2=CreateObject("ADODB.Connection")
conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended properties=Excel 5.0;Data Source="&Server.MapPath("uploadfile/excel/"&ExName) '要导的EXCEL表

errstr=""
errflag=0
set rs = conn2.execute("SELECT * FROM ["&ExTName&"$]")
if not rs.eof then
	ni=0
	do while not rs.eof
	if rs(0)<>"" then
		ni=ni+1
		if checkifnum(rs(4))=0 then
			errstr=errstr&"\n\n第"&ni&"项物资 "&rs(0)&"，规格 "&rs(1)&"，数量不正确！"
			errflag=1
		end if
	end if
	rs.movenext
	loop
end if
rs.close

if errflag-1=0 then
	call HintAndTurn(errstr,"admin_tender_add_excelin2.asp?djh="&djh)
	response.end
end if

set rs = conn2.execute("SELECT * FROM ["&ExTName&"$]")
if not rs.eof then
	ni=0
	do while not rs.eof
		if rs(0)<>"" and rs(4)<>"" then
			ni=ni+1
			set rs2=server.CreateObject("adodb.recordset")
			rs2.Open "select * from tender", conn, 1, 3
			rs2.AddNew
			rs2("djh")=djh
			rs2("material")=rs(0)
			rs2("material_guige")=rs(1)
			rs2("material_caizhi")=rs(2)
			rs2("material_danwei")=rs(3)
			rs2("material_shuliang")=rs(4)
			rs2("addman")=session("user_id")
			rs2("addtime")=now()
			rs2("finalcompany")=0
			rs2("statue")=0
			rs2("ifzu")=0
			rs2("ifdel")=0
			rs2.Update
			rs2.close
			
			If Err = 0 Then
				Response.Write "<font color='ff0000'>第"&ni&"项 导入物资 "&rs(0)&"，规格 "&rs(1)&" 成功!</font><Br>"
			Else
				Response.Write "<font color='0000ff'>第"&ni&"项 导入失败!</font>'"&Err&"'<Br>"
			End If
	
			Response.Flush
		end if
	rs.movenext
	loop
end if
rs.Close
Set rs = nothing
conn.close
set conn = nothing
conn2.close
set conn2 = Nothing

'Response.Write "<a onclick='javascript:window.close(1)'><font color='ff0000'>完闭!</font></a>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='javascript:window.location.href=""admin_tender_add_excelin1.asp""'><font color='ff0000'>返回!</font></a>"

else
response.redirect "erro.asp"
response.end
end if%>