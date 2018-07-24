<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("licq.png","系统设置")

if request("action")="saveadd" then
	nowsysname=RemoveHTML(ReplaceQM(request("sysname")))
	nowsysuser=RemoveHTML(ReplaceQM(request("sysuser")))
	nowsysmanagercontact=RemoveHTML(ReplaceQM(request("sysmanagercontact")))
	nowsysmanager=RemoveHTML(ReplaceQM(request("sysmanager")))
	nowsyscontactphone=RemoveHTML(ReplaceQM(request("syscontactphone")))
	nowgonggao=RemoveHTML(ReplaceQM(request("gonggao")))
	nowmaxrecord=checkifnum(request("maxrecord"))
	nowmaxrecord1=checkifnum(request("maxrecord1"))
	nowmaxcfile=checkifnum(request("maxcfile"))
	nowmaxccontract=checkifnum(request("maxccontract"))
	nowmaxsubtender=checkifnum(request("maxsubtender"))
	nowModalwidth=checkifnum(request("Modalwidth"))
	nowModalheight=checkifnum(request("Modalheight"))
	nowcanreg=checkifnum(request("canreg"))
	nowusercheck=checkifnum(request("usercheck"))
	nowuploadlimit=checkifnum(request("uploadlimit"))
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from config where id=1"
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("sysname")=nowsysname
		rs("sysuser")=nowsysuser
		rs("sysmanagercontact")=nowsysmanagercontact
		rs("syscontactphone")=nowsyscontactphone
		rs("sysmanager")=nowsysmanager
		rs("gonggao")=nowgonggao
		rs("maxrecord")=nowmaxrecord
		rs("maxrecord1")=nowmaxrecord1
		rs("maxcfile")=nowmaxcfile
		rs("maxccontract")=nowmaxccontract
		rs("maxsubtender")=nowmaxsubtender
		rs("Modalwidth")=nowModalwidth
		rs("Modalheight")=nowModalheight
		rs("canreg")=nowcanreg
		rs("usercheck")=nowusercheck
		rs("uploadlimit")=nowuploadlimit
		rs.update
	end if
	rs.close
	
	call HintAndTurn("系统设置成功！","admin_config.asp")
end if
%>
<script language="javascript">
function checkenter(){
	if (document.form.sysname.value == ""){
		alert("请输入系统名称！");
		form.sysname.focus();
		return false;
	}
	if (document.form.sysuser.value == ""){
		alert("请输入系统使用单位！");
		form.sysuser.focus();
		return false;
	}
}
</script>
<form name='form'>
<input type="hidden" name="action" value="saveadd">
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr>
	<td class="list_header_required" width="15%">系统名称</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='sysname' value='<%=sysname%>' size='40'><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">使用单位</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='sysuser' value='<%=sysuser%>' size='40'><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">联系方式</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='sysmanagercontact' value='<%=sysmanagercontact%>' size='80'></td>
</tr>
<tr>
	<td class="list_header_required">联系电话</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='syscontactphone' value='<%=syscontactphone%>' size='80'></td>
</tr>
<tr>
	<td class="list_header_required">负责人</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='sysmanager' value='<%=sysmanager%>' size='80'></td>
</tr>
<tr>
	<td class="list_header_required">跑马灯公告</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='gonggao' value='<%=gonggao%>' size='80' ></td>
</tr>
<tr>
	<td class="list_header_required">配置页面显示记录条数</td>
	<td class="list_required"><select name='maxrecord1' class='form_combo_normal'>
	<%for i=5 to 30
		response.write "<option value="&i
		if i-maxrecord1=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>
</tr>
<tr>
	<td class="list_header_required">信息页面显示记录条数</td>
	<td class="list_required"><select name='maxrecord' class='form_combo_normal'>
	<%for i=5 to 30
		response.write "<option value="&i
		if i-maxrecord=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>
</tr>
<tr>
	<td class="list_header_required">企业资质上传数量上限</td>
	<td class="list_required"><select name='maxcfile' class='form_combo_normal'>
	<%for i=1 to 20
		response.write "<option value="&i
		if i-maxcfile=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>
</tr>
<tr>
	<td class="list_header_required">企业合同上传数量上限</td>
	<td class="list_required"><select name='maxccontract' class='form_combo_normal'>
	<%for i=1 to 20
		response.write "<option value="&i
		if i-maxccontract=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>
</tr>
<tr>
	<td class="list_header_required">一次竞价信息中竞价物资种类上限</td>
	<td class="list_required"><select name='maxsubtender' class='form_combo_normal'>
	<%for i=50 to 500 step 50
	
		response.write "<option value="&i
		if i-maxsubtender=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>
</tr>
<tr>
	<td class="list_header_required">弹出信息框大小</td>
	<td class="list_required">宽：<input type='text' class='form_textbox_normal' name='Modalwidth' value='<%=Modalwidth%>' size='10' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
	高：<input type='text' class='form_textbox_normal' name='Modalheight' value='<%=Modalheight%>' size='10' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
</tr>
<tr>
	<td class="list_header_required">允许供应商注册</td>
	<td class="list_required">是<input type="radio" class='form_radio_normal' name="canreg" id="canreg" value="1"<%if canreg-1=0 then response.write " checked"%>> 否<input type="radio" class='form_radio_normal' name="canreg" id="canreg" value="0"<%if canreg=0 then response.write " checked"%>></td>
</tr>
<tr>
	<td class="list_header_required">供应商注册后需审核</td>
	<td class="list_required">是<input type="radio" class='form_radio_normal' name="usercheck" id="usercheck" value="0"<%if usercheck=0 then response.write " checked"%>> 否<input type="radio" class='form_radio_normal' name="usercheck" id="usercheck" value="1"<%if usercheck-1=0 then response.write " checked"%>></td>
</tr>
<tr>
	<td class="list_header_required">上传文件上限</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='uploadlimit' value='<%=uploadlimit%>' size='10' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> M</td>
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