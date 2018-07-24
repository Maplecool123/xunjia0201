<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("identity.png","数据中心")

Dim ZC_DATABASE_PATH
'数据库路径
ZC_DATABASE_PATH="shujuku/#jingbiaodata#.mdb"

data_array= Split(ZC_DATABASE_PATH,"/")

Dim action
action=trim(request("action"))
Dim dbpath,bkfolder,bkdbname,fso,fso1

Select Case action
Case ""
Call chushihua()
Case "CompressData" '压缩数据
Dim tmprs
dim allarticle
dim Maxid
dim topic,username,dateandtime,body
call CompressData()
case "BackupData" '备份数据
if request("act")="Backup" Then
call updata()
else
call BackupData()
end If
case "RestoreData" '还原数据
dim backpath
if request("act")="Restore" Then
Dbpath=request.form("Dbpath")
backpath=request.form("backpath")
if dbpath="" Then
response.write "Please input your database whole Name" 
else
Dbpath=server.mappath(Dbpath)
end If
backpath=server.mappath(backpath)

Set Fso=server.CreateObject("scripting.filesystemobject")
if fso.fileexists(dbpath) Then 
fso.copyfile Dbpath,Backpath
response.write "数据库还原成功！<br><br><a href=""data_s.asp"">返回</a>"
  
else
response.write "没有找到您所需的数据库！" 
response.write "<br>输入路径错误，请确认后重新输入！<br><br><a href=""data_s.asp"">返回</a>"
end If
else
call RestoreData()
end If
Case "SpaceSize" '系统空间占用
call SpaceSize()
Case "deletebackup"
Dim dbname
dbpath=Request.QueryString("dbpath")
dbname=Request.QueryString("dbname")
dbpath=Server.MapPath(dbpath)
dbpath=dbpath &"\"&dbname
set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
fso.DeleteFile(DBPath)
Set fso = nothing
response.write "<br>您备份的数据库已经" & dbpath &"被成功删除<br><br><a href=""data_s.asp"">返回</a>"
Else
response.write dbpath 
response.write "<br>输入路径错误，请确认后重新输入！<br><br><a href=""data_s.asp"">返回</a>"
End If
Case Else
End Select

%>
</div>
<%

Sub chushihua()
%>
<div align=center>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td height="25" colspan="4">数据中心</td>
</tr>
<form id="edit">
<tr height="25" align="center" class="grid_odd">
	<td><a href="?action=CompressData"><b>[压缩数据库]</b></a></td>
	<td><a href="?action=BackupData"><b>[备份数据库]</b></a></td>
	<td><a href="?action=RestoreData"><b>[还原数据库]</b></a></td>
	<td><a href="?action=SpaceSize"><b>[系统占用空间]</b></a></td>
</tr>
</form>
</table>
</div>
<%end sub%>

<%
'====================系统空间占用=======================
Sub SpaceSize()
On Error Resume Next
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td height="25" colspan="4">系统空间察看</td>
</tr>
<form id="edit">
<tr height="25" align="center" class="grid_odd">
	<td><b>数据库：<%showSpaceinfo("../"&data_array(1)&"")%></b></td>
</tr>
<tr height="25" align="center" class="grid_odd">
	<td><b>备份数据库：<%showSpaceinfo("databackup")%></b></td>
</tr>
<tr height="25" align="center" class="grid_odd">
	<td><b>系统共计：<%showSpaceinfo("/")%></b></td>
</tr>
<tr height="25" align="center" class="grid_odd">
	<td><input class='form_submit' type="button" name="back" value=" 返回 " onClick="window.location.href='data_s.asp'"></td>
</tr>
</form>
</table>
<%
End Sub

Sub ShowSpaceInfo(drvpath)
dim fso,d,size,showsize
set fso=server.CreateObject("scripting.filesystemobject") 
drvpath=server.mappath(drvpath) 
set d=fso.getfolder(drvpath) 
size=d.size
showsize=size & " Byte" 
if size>1024 Then
size=(Size/1024)
showsize=size & " KB"
end If
if size>1024 Then
size=(size/1024)
showsize=formatnumber(size,2) & " MB" 
end If
if size>1024 Then
size=(size/1024)
showsize=formatnumber(size,2) & " GB" 
end If 
response.write "<font face=verdana>" & showsize & "</font>"
End Sub 

Sub RestoreData()
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td height="25" colspan="4">还原数据库</td>
</tr>
<form id="edit" method="post" action="?action=RestoreData&act=Restore">
<tr height="25" class="grid_odd">
	<td width="50%" align="right"><b>还原的路径（相对路径）：</b></td>
	<td width="50%" align="left"><input type="text" size="30" name="DBpath" value="DataBackup\<%=replace(Date(),"/","-")%>_Bak.mdb"></td>
</tr>
<tr height="25" class="grid_odd">
	<td align="right"><b>还原后的路径（相对路径）：</b></td>
	<td align="left"><input type="text" size="30" name="backpath" value="<%=ZC_DATABASE_PATH%>"></td>
</tr>
<tr height="25" class="grid_odd" align="center">
	<td colspan="2"><input type="submit" value="开始还原" class="form_button"></td>
</tr>
<tr height="25" align="center" class="grid_odd">
	<td colspan="2"><input class='form_submit' type="button" name="back" value=" 返回 " onClick="window.location.href='data_s.asp'"></td>
</tr>
</form>
</table>
<%
End Sub

Sub updata()
Dbpath=request.form("Dbpath")
Dbpath=server.mappath(Dbpath)
bkfolder=request.form("bkfolder")
bkdbname=request.form("bkdbname")
Set Fso=server.CreateObject("scripting.filesystemobject")
if fso.fileexists(dbpath) Then
	If CheckDir(bkfolder) = True Then
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
	else
		MakeNewsDir bkfolder
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
	end If
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td height="25" colspan="4">备份数据库</td>
</tr>
<tr height="25" class="grid_odd">
	<td width="50%" align="center"><b>已经成功备份，您的数据库路径：<%=bkfolder%>\<%=bkdbname%></b></td>
</tr>
<%response.write "<tr height='25' class='grid_odd'>"
response.write "<td align='center'><b>"
response.write "点击此处下载数据库：<font color=red><a href="""& ZC_BLOG_HOST & request.form("bkfolder") & "/" & bkdbname &""">" & ZC_BLOG_HOST & request.form("bkfolder") & "/" & bkdbname
response.write "</font></td></tr>"
response.write "<tr height='25' class='grid_odd'>"
response.write "<td align='center'><b>"
response.write "<a href=""data_s.asp?action=deletebackup&dbpath="&request.form("bkfolder") &"&dbname=" & bkdbname &""">当您下载完毕后可点击此处删除备份数据库！</a>"
response.write "</td></tr>"
response.write "<tr height='25' class='grid_odd'>"
response.write "<td align='center'><b>"
response.write "<input class='form_submit' type='button' name='back' value=' 返回 ' onClick='window.location.href=""data_s.asp"";'></td></tr>"
%>
</table>
<%Else
	response.write "Error ,找不到文件<br>"
End If
Set fso = nothing
End Sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
folderpath=Server.MapPath(".")&"\"&folderpath
Set fso1 = CreateObject("Scripting.FileSystemObject")
If fso1.FolderExists(FolderPath) Then
'存在
CheckDir = True
Else
'不存在
CheckDir = False
End If
Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
dim f
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set f = fso1.CreateFolder(foldername)
MakeNewsDir = True
Set fso1 = nothing
End Function
Sub BackupData()
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td height="25" colspan="4">备份数据库</td>
</tr>
<form id="edit" method="post" action="?action=BackupData&act=Backup">
<tr height="25" class="grid_odd">
	<td width="50%" align="right"><b>当前数据库路径（相对路径）：</b></td>
	<td width="50%" align="left"><input type="text" size="15" name="DBpath" value="<%=ZC_DATABASE_PATH%>"></td>
</tr>
<tr height="25" class="grid_odd">
	<td align="right"><b>备份数据库路径（相对路径）：</b></td>
	<td align="left"><input type="text" size="15" name="bkfolder" value="Databackup">该目录若不存在，系统将自动建立</td>
</tr>
<tr height="25" class="grid_odd">
	<td align="right"><b>备份后数据库名称</b></td>
	<td align="left"><input type="text" size="20" name="bkDBname" value="<%=replace(Date(),"/","-")%>_bak.mdb">按日期自动命名</td>
</tr>
<tr height="25" class="grid_odd">
	<td align="center" colspan="2"><b>如果备份文件不存在将建立；如果存在将覆盖</b></td>
</tr>
<tr height="25" class="grid_odd">
	<td align="center" colspan="2"><input type="submit" value="开始备份" class="form_button"></td>
</tr>
<tr height="25" align="center" class="grid_odd">
	<td colspan="2"><input class='form_submit' type="button" name="back" value=" 返回 " onClick="window.location.href='data_s.asp'"></td>
</tr>
</form>
</table>
<%
End Sub

Sub CompressData()
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td height="25" colspan="4">压缩数据库</td>
</tr>
<form id="edit" action="?action=CompressData" method="post">
<tr height="25" class="grid_odd">
	<td width="50%" align="right"><b>压缩数据库的路径：</b></td>
	<td width="50%" align="left"><input type="text" name="dbpath" value="<%=ZC_DATABASE_PATH%>"></td>
</tr>
<tr height="25" class="grid_odd">
	<td align="center" colspan="2"><input type="submit" value="开始压缩" class="form_button"></td>
</tr>
<tr height="25" align="center" class="grid_odd">
	<td colspan="2"><input class='form_submit' type="button" name="back" value=" 返回 " onClick="window.location.href='data_s.asp'"></td>
</tr>
</form>
</table>
<%
Dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
response.write(CompactDB(dbpath,boolIs97))
End If

End Sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = Left(dbPath,InStrRev(DBPath,"\"))
'response.write strDBPath
'response.end
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
fso.CopyFile dbpath,strDBPath & "temp.mdb"
Set Engine = CreateObject("JRO.JetEngine")

If boolIs97 = "True" Then
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;User ID=admin;Password=;Jet OLEDB:Database Password=;", _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;User ID=admin;Password=;Jet OLEDB:Database Password=;" _
& "Jet OLEDB:Engine Type=" & JET_3X
Else
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;User ID=admin;Password=;Jet OLEDB:Database Password=;", _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;User ID=admin;Password=;Jet OLEDB:Database Password=;"
End If

fso.CopyFile strDBPath & "temp1.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
fso.DeleteFile(strDBPath & "temp1.mdb")
Set fso = nothing
Set Engine = nothing

CompactDB = "您的数据库" & dbpath & "已经成功被压缩！" & vbCrLf

  
Else
CompactDB = "<br>您输入的路径有误，请确认后重新输入！" & vbCrLf
End If

End Function

'////////////////////end////////////////////////

else
response.redirect "erro.asp"
response.end
end if%>