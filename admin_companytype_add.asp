<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("gpa.png","添加新公司类型")

nowcompanytypename=RemoveHTML(ReplaceQM(request("companytypename")))
nowcompanytypeorder=checkifnum(request("companytypeorder"))
	
if request("action")="saveadd" then
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from companytype where companytypename='"&nowcompanytypename&"'"
	rs.open sql,conn,1,3
	if not rs.eof then
		call HintAndBack("您输入的公司类型名称已经存在,请重新输入！",1)
	else
		rs.addnew
		rs("companytypename")=nowcompanytypename
		rs("companytypeorder")=nowcompanytypeorder
		rs.update
	end if
	rs.close
	
	call HintAndTurn("添加新公司类型！","admin_companytype.asp")
end if
%>
<script language="javascript">
function checkenter(){
	if (document.form.companytypename.value == ""){
		alert("请填写公司类型名称！");
		form.companytypename.focus();
		return false;
	}
	if (document.form.companytypeorder.value == ""){
		alert("请填写公司类型排序！");
		form.companytypeorder.focus();
		return false;
	}
}
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form'>
<input type="hidden" name="action" value="saveadd">
<tr>
	<td class="list_header_required" width="15%">公司类型名称</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companytypename' id='companytypename' value='' size='20'><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">公司类型排序</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companytypeorder' id='companytypeorder' value='' size='20' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"><font color="#FF0000">*</font></td>
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