<!-- #include file="datatype.inc" -->
<%
if closesystem-1=0 then
	response.write "<center><font size=5><strong>系统维护中，请稍后访问！</strong></font></center>"
	response.end
else
	set conn=server.createobject("adodb.connection")
	If IsSqlDataBase=0 Then
	mypath=server.mappath("shujuku/#jingbiaodata#.mdb")
	'ConnStr="driver={microsoft access driver (*.mdb)};dbq="&mypath
	ConnStr="provider=microsoft.jet.oledb.4.0;data source=" & mypath
	else
	'ConnStr = "Driver={SQL Server};SERVER=476D5C85323C48D;UID=sa;PWD=19830315ld;DATABASE=kcmanage10"
	ConnStr = "Driver={SQL Server};SERVER=F12920534BB34D6;UID=sa;PWD=w12345;DATABASE=kcmanage10"
	'SqlNowString="GetDate()"
	end if
	conn.open ConnStr
end if
%>                                                                                                                                                                                                                                                                
