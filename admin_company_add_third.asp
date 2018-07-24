<!--#include file="Index_top.asp" -->
<!--#include file="inc/uploadcontrol.asp" -->
<%
if session("iflogin")-2=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","添加企业第三步：上传企业以往合同")
%>
<script type="text/javascript">
function firm(firmtxt,firmno){
if(confirm(firmtxt)){
	eval("document.form.companycontract"+ firmno +".value=''");
	eval("document.all.upid"+ firmno +".style.display='none';");
}
}
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form' method="post" action="upprocess_forth.asp" enctype="multipart/form-data">
<input type="hidden" name="action" value="uploadfile">
<input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="gotourl" value="admin_company">
<input type="hidden" name="companyrealname" value="<%=request("companyrealname")%>">
<tr>
	<td class="list_header_required" width="15%">上传合同<br><font color="#FF0000">请上传公司（企业）以往合同盖章扫描件至少3份</font></td>
	<td class="list_required" colspan="2">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr valign="middle">
			<td class="list_required">上传文件大小不可超过：<%=uploadlimit%> M，可上传格式为：<%=Forumuploadall%></td>
		</tr>
		<%for nnn=1 to maxcfile%>
		<tr valign="middle" id="upid<%=nnn%>"<%if nnn>3 then%> style="display:none;"<%end if%>>
			<td class="list_required">文件<%=right("0"&nnn,2)%>：
			<input type="file" name="companycontract<%=nnn%>" style="width:400" class="form_button">
			<%if nnn<>maxcfile then%><input type="button" value="添加另一个文件" onClick="upid<%=(nnn+1)%>.style.display='';" class="form_button">&nbsp;&nbsp;<%end if%>
			<%if nnn<>1 then%><input type="button" value="清  除" onClick="firm('确定清除第 <%=nnn%> 条么？',<%=nnn%>);" class="form_submit"><%end if%></td>
		</tr>
		<%next%>
		</table>
	</td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="2"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 添加 '>
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-2);"></td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->