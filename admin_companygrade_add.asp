<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("gpa.png","添加新供应商资质级别")

nowcompanygradename=RemoveHTML(ReplaceQM(request("companygradename")))
nowcompanygradeorder=checkifnum(request("companygradeorder"))
	
if request("action")="saveadd" then
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from companygrade where companygradename='"&nowcompanygradename&"'"
	rs.open sql,conn,1,3
	if not rs.eof then
		call HintAndBack("您输入的供应商级别名称已经存在,请重新输入！",1)
	else
		rs.addnew
		rs("companygradename")=nowcompanygradename
		rs("companygradeorder")=nowcompanygradeorder
		rs.update
	end if
	rs.close
	
	call HintAndTurn("添加新供应商级别！","admin_companygrade.asp")
end if
%>
<script language="javascript">
function checkenter(){
	if (document.form.companygradename.value == ""){
		alert("请填写供应商级别名称！");
		form.companygradename.focus();
		return false;
	}
	if (document.form.companygradeorder.value == ""){
		alert("请填写级别高低数值！");
		form.companygradeorder.focus();
		return false;
	}
}
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form'>
<input type="hidden" name="action" value="saveadd">
<tr>
	<td class="list_header_required" width="15%">供应商级别名称</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companygradename' id='companygradename' value='' size='20'><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">级别高低数值</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companygradeorder' id='companygradeorder' value='' size='20' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command"><input class='form_submit' type='submit' name='cmd' value=' 添加 ' onClick="return checkenter()">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->