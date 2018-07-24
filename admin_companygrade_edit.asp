<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("gpa.png","修改供应商资质级别")

gourl="checkstyle="&request("checkstyle")&"&keyword1="&request("keyword1")&"&page="&request("page")

nowcompanygradename=RemoveHTML(ReplaceQM(request("companygradename")))
nowcompanygradeorder=checkifnum(request("companygradeorder"))
nowoldcompanygradeorder=checkifnum(request("oldcompanygradeorder"))
	
if request("action")="saveadd" then
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from companygrade where companygradename='"&nowcompanygradename&"' and id<>"&request("id")&""
	rs.open sql,conn,1,3
	if not rs.eof then
		call HintAndBack("您输入供应商级别名称已经存在,请重新输入！",1)
	end if
	rs.close
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from companygrade where id="&request("id")&""
	rs.open sql,conn,1,3
	rs("companygradename")=nowcompanygradename
	rs("companygradeorder")=nowcompanygradeorder
	rs.update
	rs.close
	
	if nowoldcompanygradeorder-nowcompanygradeorder<>0 then
		conn.execute("update tender set tendergrade="&nowcompanygradeorder&" where tendergradeid="&request("id")&"")
	end if
	
	call HintAndTurn("修改供应商级别成功！","admin_companygrade.asp?"&gourl&"")
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
<%
if request("action")="" then
	set rsm=conn.execute("select * from companygrade where id="&request("id")&"")
	if not rsm.eof then
		companygradename=rsm("companygradename")
		companygradeorder=rsm("companygradeorder")
	end if
	rsm.close
end if%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form'>
<input type="hidden" name="action" value="saveadd">
<input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="checkstyle" value="<%=request("checkstyle")%>">
<input type="hidden" name="keyword1" value="<%=request("keyword1")%>">
<input type="hidden" name="page" value="<%=request("page")%>">
<input type="hidden" name="oldcompanygradeorder" value="<%=companygradeorder%>">
<tr>
	<td class="list_header_required" width="15%">供应商级别名称</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companygradename' value='<%=companygradename%>' size='20'><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">级别高低数值</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companygradeorder' value='<%=companygradeorder%>' size='20' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command"><input class='form_submit' type='submit' name='cmd' value=' 修改 ' onClick="return checkenter()">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->