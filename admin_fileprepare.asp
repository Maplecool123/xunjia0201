<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("licq.png","供应商文件格式要求")

if request.form("action")="saveadd" then
	nowfileprepare=request.form("fileprepare")
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from config where id=1"
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("fileprepare")=nowfileprepare
		rs.update
	end if
	rs.close
	
	call HintAndTurn("供应商注册须知设置成功！","admin_fileprepare.asp")
end if
%>
<script type="text/javascript" charset="utf-8" src="edit/kindeditor.js"></script>
<script type="text/javascript">
KE.show({
id : 'fileprepare',
cssPath : 'edit/index.css'
});

function checkenter(){
	if (document.form.fileprepare.value == ""){
		alert("请输入供应商文件格式要求！");
		return false;
	}
}
</script>
<form name='form' method="post">
<input type="hidden" name="action" value="saveadd">
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr>
	<td class="list_header_required" width="15%">供应商文件格式要求</td>
	<td class="list_required"><textarea id="fileprepare" name="fileprepare" style="width:620px;height:400px;visibility:hidden;"><%=fileprepare%></textarea></td>
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