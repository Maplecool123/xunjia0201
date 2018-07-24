<!--#include file="conn.asp"-->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"   
response.charset="gb2312"           

companyrealname=request("companyrealname")
Set rs=conn.execute("Select * from [company] where companyrealname='"&companyrealname&"'")
if not rs.eof then
    response.write("<img src='images/err.gif' border='0' width='18' height='18'> <font color='#ff0000'>该企业已经注册，请勿重复注册！</font>")
else
    response.write("<img src='images/succ.gif' border='0' width='18' height='18'> <font color='#0000ff'>该企业尚未注册，请继续！</font>")
end if
rs.close
set rs=nothing
conn.close
Set conn=nothing
%>
