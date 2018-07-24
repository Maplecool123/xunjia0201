<!--#include file="Index_top.asp" -->
<!--#include file="inc/uploadcontrol.asp" -->
<%
if session("iflogin")-2=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","添加企业第二步：上传企业资质材料")
%>
<script type="text/javascript">
function checkenter(){
	if (document.form.companyfile1.value == ""){
		alert("请上传竞价采购承诺书！");
		form.companyfile1.focus();
		return false;
	}
	if (document.form.companyfile2.value == ""){
		alert("请上传企业法人营业执照！");
		form.companyfile2.focus();
		return false;
	}
	if (document.form.companyfile3.value == ""){
		alert("请上传税务登记证！");
		form.companyfile3.focus();
		return false;
	}
	if (document.form.companyfile4.value == ""){
		alert("请上传组织机构代码证！");
		form.companyfile4.focus();
		return false;
	}
}
function firm(firmtxt,firmno){
if(confirm(firmtxt)){
	eval("document.form.companyfile"+ firmno +".value=''");
	eval("document.all.upid"+ firmno +".style.display='none';");
}
}
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form' method="post" action="upprocess_third.asp" enctype="multipart/form-data">
<input type="hidden" name="action" value="uploadfile">
<input type="hidden" name="gotourl" value="admin_company_add_third">
<input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="companyrealname" value="<%=request("companyrealname")%>">
<tr>
	<td class="list_header_required" width="15%">上传资质材料<br><font color="#FF0000">请上传真实的公司（企业）资质盖章扫描件</font></td>
	<td class="list_required" colspan="2">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr valign="middle">
			<td class="list_required">上传文件大小不可超过：<%=uploadlimit%> M，可上传格式为：<%=Forumuploadall%></td>
		</tr>		<tr valign="middle">
			<td class="list_required">竞价采购承诺书：　　　
			<input type="file" name="companyfile1" style="width:400" class="form_button"></td>
		</tr>
		<tr valign="middle">
			<td class="list_required">企业法人营业执照：　　
			<input type="file" name="companyfile2" style="width:400" class="form_button"></td>
		</tr>
		<tr valign="middle">
			<td class="list_required">税务登记证：　　　　　
			<input type="file" name="companyfile3" style="width:400" class="form_button"></td>
		</tr>
		<tr valign="middle">
			<td class="list_required">组织机构代码证：　　　
			<input type="file" name="companyfile4" style="width:400" class="form_button">
			<%if maxcfile>4 then%>
			<input type="button" value="添加另一个文件" onClick="upid5.style.display='';" class="form_button">&nbsp;&nbsp;
			<%end if%></td>
		</tr>
		<%for nnn=5 to maxcfile%>
		<tr valign="middle" id="upid<%=nnn%>"<%if nnn<>1 then%> style="display:none;"<%end if%>>
			<td class="list_required">文件<%=right("0"&nnn,2)%>：
			<input type="file" name="companyfile<%=nnn%>" style="width:400" class="form_button">
			<%if nnn<>maxcfile then%><input type="button" value="添加另一个文件" onClick="upid<%=(nnn+1)%>.style.display='';" class="form_button">&nbsp;&nbsp;<%end if%>
			<%if nnn<>1 then%><input type="button" value="清  除" onClick="firm('确定清除第 <%=nnn%> 条么？',<%=nnn%>);" class="form_submit"><%end if%></td>
		</tr>
		<%next%>
		</table>
	</td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="2"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 添加 ' onClick="return checkenter()">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-2);"></td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->