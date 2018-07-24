<!--#include file="Index_top_nologin.asp" -->
<title><%=sysname%> - 企业注册第三步</title>
<!--#include file="inc/Md5.asp" -->
<%
call WhereTable("identity.png","企业注册第三步：填写基本信息")

if request.form("action")="saveadd" then
	nowcompanyclass=request.form("companyclass")
	nowcompanytype=request.form("companytype")
	nowcompanyrealname=RemoveHTML(ReplaceQM(request.form("companyrealname")))
	nowcompanyname=RemoveHTML(ReplaceQM(request.form("companyname")))
	nowpwd=RemoveHTML(ReplaceQM(request.form("pwd")))
	nowcompanyorgno=RemoveHTML(ReplaceQM(request.form("companyorgno")))
	nowcompanytaxno=RemoveHTML(ReplaceQM(request.form("companytaxno")))
	nowcompanymanager=RemoveHTML(ReplaceQM(request.form("companymanager")))
	nowcompanycontactman=RemoveHTML(ReplaceQM(request.form("companycontactman")))
	nowcompanycontacttel=RemoveHTML(ReplaceQM(request.form("companycontacttel")))
	nowcompanyaddress=RemoveHTML(ReplaceQM(request.form("companyaddress")))
	nowcompanyemail=request.form("companyemail")
	nowcompanymajor=request.form("companymajor")
	nowcompanybeizhu=request.form("companybeizhu")
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from company where companyrealname='"&nowcompanyrealname&"' or companyname='"&nowcompanyname&"'"
	rs.open sql,conn,1,3
	if not rs.eof then
		call HintAndBack("您输入的企业名称或登录用户名已经存在,请重新输入！",1)
	else
		rs.addnew
		rs("companytype")=nowcompanytype
		rs("companyrealname")=nowcompanyrealname
		rs("companyname")=nowcompanyname
		if nowpwd="" then
			rs("pwd")=md5("666666")
		else
			rs("pwd")=md5(nowpwd)
		end if
		rs("companyorgno")=nowcompanyorgno
		rs("companytaxno")=nowcompanytaxno
		rs("companymanager")=nowcompanymanager
		rs("companycontactman")=nowcompanycontactman
		rs("companycontacttel")=nowcompanycontacttel
		rs("companyaddress")=nowcompanyaddress
		rs("companyemail")=nowcompanyemail
		rs("companymajor")=nowcompanymajor
		rs("companybeizhu")=nowcompanybeizhu
		rs("companyregtime")=now()
		rs("companyclass")=nowcompanyclass
		rs("companygrade")=0
		rs("statue")=usercheck
		rs("regstatue")=1
		rs.update
	end if
	rs.close
	
	call HintAndTurn("基本信息设置成功，下一步上传企业资质材料！","company_reg_forth.asp?id="&shownamestrnull("id","company","companyname",nowcompanyname,"")&"&companyrealname="&nowcompanyrealname)
end if
%>
<script type="text/javascript" charset="utf-8" src="edit/kindeditor.js"></script>
<script type="text/javascript">
KE.show({
id : 'companybeizhu',
cs sPath : 'edit/index.css'
});

function checkenter(){
	if (document.form.companyrealname.value == ""){
		alert("请填写企业名称！");
		form.companyrealname.focus();
		return false;
	}
	if (document.form.companyname.value == ""){
		alert("请填写登录用户名！");
		form.companyname.focus();
		return false;
	}
	if (document.form.companyorgno.value == ""){
		alert("请填写组织机构代码证号！");
		form.companyorgno.focus();
		return false;
	}
	if (document.form.companytaxno.value == ""){
		alert("请填写税号！");
		form.companytaxno.focus();
		return false;
	}
	if (document.form.companymanager.value == ""){
		alert("请填写法定代表人！");
		form.companymanager.focus();
		return false;
	}
	if (document.form.companycontactman.value == ""){
		alert("请填写联系人姓名！");
		form.companycontactman.focus();
		return false;
	}
	if (document.form.companycontacttel.value == ""){
		alert("请填写联系电话！");
		form.companycontacttel.focus();
		return false;
	}
	if (document.form.companyaddress.value == ""){
		alert("请填写公司地址！");
		form.companyaddress.focus();
		return false;
	}
	if (document.form.companyemail.value == ""){
		alert("请填写公司电子邮件！");
		form.companyemail.focus();
		return false;
	}
}

var xmlHttp = false;
try {
  xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
} catch (e) {
  try {
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  } catch (e2) {
    xmlHttp = false;
  }
}
if (!xmlHttp && typeof XMLHttpRequest != 'undefined') {
  xmlHttp = new XMLHttpRequest();
}

function callServer1() {
  var p_name = document.getElementById("companyrealname").value;
  if ((p_name == null) || (p_name == "")) return;
  var url = "Checkcompany1.asp?companyrealname=" + escape(p_name);
  xmlHttp.open("GET", url, true);
  xmlHttp.onreadystatechange = updatePage1;
  xmlHttp.send(null);  
}

function updatePage1() {
  if (xmlHttp.readyState < 4) {
    test1.innerHTML="正在查询该用户名是否已存在";
  }
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText;
    test1.innerHTML=response;
  }
  if (response.indexOf("err")>=0){
    document.form.cmd.disabled=true;
  }
  else
  {
  	if(document.all.test2.innerHTML=="" || document.all.test2.innerHTML.indexOf("succ")>=0)
	{
    	document.form.cmd.disabled=false;
	}
	else
	{
    	document.form.cmd.disabled=true;
	}
  }
}

function callServer2() {
  var p_name = document.getElementById("companyname").value;
  if ((p_name == null) || (p_name == "")) return;
  var url = "Checkcompany2.asp?companyname=" + escape(p_name);
  xmlHttp.open("GET", url, true);
  xmlHttp.onreadystatechange = updatePage2;
  xmlHttp.send(null);  
}

function updatePage2() {
  if (xmlHttp.readyState < 4) {
    test2.innerHTML="正在查询该用户名是否已存在";
  }
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText;
    test2.innerHTML=response;
  }
  if (response.indexOf("err")>=0){
    document.form.cmd.disabled=true;
  }
  else
  {
  	if(document.all.test1.innerHTML=="" || document.all.test1.innerHTML.indexOf("succ")>=0)
	{
    	document.form.cmd.disabled=false;
	}
	else
	{
    	document.form.cmd.disabled=true;
	}
  }
}
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form' method="post">
<input type="hidden" name="action" value="saveadd">
<tr>
	<td class="list_header_required">企业类型</td>
	<td class="list_required" colspan="2"><select class='form_combo_normal' name='companytype'>
	<%set rsa=conn.execute("select * from companytype order by companytypeorder")
	if not rsa.eof then
		do while not rsa.eof
		response.write "<option value="&rsa("id")&">"&rsa("companytypename")&"</option>"
		rsa.movenext
		loop
	end if
	rsa.close%>
	</select><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required" width="15%">企业名称</td>
	<td class="list_required" width="45%"><input type='text' class='form_textbox_normal' name='companyrealname' id='companyrealname' value='' onKeyUp="callServer1();" style="width:300px;"><font color="#FF0000">*</font></td>
	<td class="list_required" width="40%"><span id="test1"></span></td>
</tr>
<tr>
	<td class="list_header_required">登陆用户名</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='companyname' id='companyname' value='' onKeyUp="callServer2();" style="width:200px;"><font color="#FF0000">*</font></td>
	<td class="list_required"><span id="test2"></span></td>
</tr>
<tr>
	<td class="list_header_required">组织机构代码证号</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companyorgno' value='' style="width:200px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">税号</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companytaxno' value='' style="width:200px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">登陆密码</td>
	<td class="list_required" colspan="2"><input type='password' class='form_password_required' name='pwd' value='' style="width:200px;"><font color="#666666">（不填写默认为“666666”）</font></td>
</tr>
<tr>
	<td class="list_header_required">法定代表人</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companymanager' value='' style="width:200px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">联系人姓名</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companycontactman' value='' style="width:200px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">联系电话</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companycontacttel' value='' style="width:200px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">公司地址</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companyaddress' value='' style="width:300px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">电子邮件</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companyemail' value='' style="width:200px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">主营业务</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='companymajor' value='' style="width:200px;"><font color="#666666">（请勿超过20字）</font></td>
</tr>
<tr>
	<td class="list_header_required">备注</td>
	<td class="list_required" colspan="2"><textarea id="companybeizhu" name="companybeizhu" style="width:620px;height:150px;visibility:hidden;"></textarea></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="2"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 添加 ' onClick="return checkenter()">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</form>
</table>
<!--#include file="Index_bottom.asp" -->