<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("licq.png","供应商注册须知")

if request.form("action")="saveadd" then
	nowshouldknow=request.form("shouldknow")
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from config where id=1"
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("shouldknow")=nowshouldknow
		rs.update
	end if
	rs.close
	
	call HintAndTurn("供应商注册须知设置成功！","admin_shouldknow.asp")
end if
%>
<script type="text/javascript" charset="utf-8" src="edit/kindeditor.js"></script>
<script type="text/javascript">
KE.show({
id : 'shouldknow',
cssPath : 'edit/index.css'
});

function checkenter(){
	if (document.form.shouldknow.value == ""){
		alert("请输入供应商注册须知！");
		return false;
	}
}
</script>
<form name='form' method="post">
<input type="hidden" name="action" value="saveadd">
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr>
	<td class="list_header_required" width="15%">供应商注册须知</td>
	<td class="list_required"><textarea id="shouldknow" name="shouldknow" style="width:620px;height:400px;visibility:hidden;"><%=shouldknow%></textarea></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command"><input class='form_submit' type='submit' name='cmd' value=' 确定 ' onClick="return checkenter()"></td>
</tr>
</table>
</form>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->