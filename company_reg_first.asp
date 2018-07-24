<!--#include file="Index_top_nologin.asp" -->
<title><%=sysname%> - 企业注册第一步</title>
<!--#include file="inc/Md5.asp" -->
<%
call WhereTable("identity.png","企业注册第一步：材料准备")
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all">
<form name='form'>
<tr>
	<td class="list_required"><%=fileprepare%></td>
</tr>
<tr>
	<td class="list_command" align="center"><input class='form_submit' type='button' name='cmd' id='cmd' value=' 我同意 ' onClick="window.location.href='company_reg_second.asp'">
	&nbsp;&nbsp;&nbsp;&nbsp;<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</form>
</table>
<!--#include file="Index_bottom.asp" -->