<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
'全局变量
Dim CurrentPage,sql,i,rs
'全局变量

call WhereTable("view_text.png","日志管理")

dim nowlogin_zu,nowcheckstyle,nowkeyword
currentpage=request("page")
'if currentpage="" then currentpage=0
'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

nowlogin_zu = request("login_zu")
nowlogin_dept = request("login_dept")
nowcheckstyle = request("checkstyle")
nowcheckstatue = request("checkstatue")
nowkeyword1 = RemoveHTML(ReplaceQM(request("keyword1")))

gourl="login_zu="&nowlogin_zu&"&login_dept="&nowlogin_dept&"&checkstyle="&nowcheckstyle&"&checkstatue="&nowcheckstatue&"&keyword1="&nowkeyword1&""

if request("action")="delall" then
	nowtimemm=date()
	tianshu=request("tianshu")

	conn.execute("delete from [rizi] where #"&nowtimemm&"#-logindate>"&tianshu&"")
	
	call HintAndTurn("删除"&tianshu&"天以外的日志成功！","?"&gourl&"&page="&currentpage&"")
end if

if request("action")="delmany" then
	id=replace(request("id")," ","")

	goon=0
	if replace(id,",","")<>"" then goon=1

	if goon=0 then
		call HintAndBack("请至少选中一条记录！",1)
	end if
	
	id=split(id,",")
	for i=0 to UBound(id)
		conn.execute("delete from [rizi] where id="&id(i)&"")
	next
	
	
	call HintAndTurn("删除成功！","?"&gourl&"&page="&currentpage&"")
end if
%>
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
  
function checkenter(){
	if (document.form3.tianshu.value == ""){
		alert("请填写清除日志的天数范围！");
		form3.tianshu.focus();
		return false;
	}
}
</script>
<table class='grid_search' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td width='170'>查询条件</td>
	<td width='140'>用户名或姓名</td>
	<td width='70'>查询</td>
	<td width='300'>清除日志</td>
	<td width='*'></td>
</tr>
<tr class="grid_odd" align="center">
<form name="form2">
	<td class="grid_cell"><select name="checkstatue" class='form_combo_normal' onChange="form2.submit();">
	<option value="">所有状态</option>
	<option value="登陆成功"<%if nowcheckstatue="登陆成功" then%> selected="selected"<%end if%>>登陆成功</option>
	<option value="登陆失败"<%if nowcheckstatue="登陆失败" then%> selected="selected"<%end if%>>登陆失败</option>
	</select>
	<select name="checkstyle" class='form_combo_normal' onChange="form2.submit();">
	<option value="mohu"<%if nowcheckstyle="mohu" then%> selected="selected"<%end if%>>模糊</option>
	<option value="jing"<%if nowcheckstyle="jing" then%> selected="selected"<%end if%>>精确</option>
	</select></td>
	<td class="grid_cell"><input type='text' class='form_textbox_normal' name='keyword1' value='<%=nowkeyword1%>' size='15' ></td>
	<td class="grid_cell"><input type='submit' value='查询' class='form_button'></td>
</form>
<form name="form3">
<input type="hidden" name="action" value="delall">
<input type="hidden" name="login_zu" value="<%=nowlogin_zu%>">
<input type="hidden" name="login_dept" value="<%=nowlogin_dept%>">
<input type="hidden" name="checkstyle" value="<%=nowcheckstyle%>">
<input type="hidden" name="checkstatue" value="<%=nowcheckstatue%>">
<input type="hidden" name="keyword1" value="<%=nowkeyword1%>">
	<td class="grid_cell"><input type='text' class='form_textbox_normal' name='tianshu' value='<%=tianshu%>' size='10' onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">天以外的日志
	<input type='submit' value='删除' class='form_submit' onClick="return checkenter()"></td>
</form>
	<td class="grid_cell" align="left"></td>
</tr>
</table>
<%
dim thisClass
thisClass = "grid_even"

Set Rs = server.createobject("adodb.recordset")
sql="select * from rizi where id<>0"
if nowcheckstatue<>"" then
	sql=sql&" and class='"&nowcheckstatue&"'"
end if
if nowcheckstyle="mohu" then
	if nowkeyword1<>"" then
		sql=sql&" and (username like '%"&nowkeyword1&"%' or username in (select username from login where userrealname like '%"&nowkeyword1&"%'))"
	end if
elseif nowcheckstyle="jing" then
	if nowkeyword1<>"" then
		sql=sql&" and (username = '"&nowkeyword1&"' or username in (select username from login where userrealname='"&nowkeyword1&"'))"
	end if
end if
sql=sql&" order by logindate desc"%>
<form name='form1'>
<input type="hidden" name="action" value="delmany">
<input type="hidden" name="login_zu" value="<%=nowlogin_zu%>">
<input type="hidden" name="login_dept" value="<%=nowlogin_dept%>">
<input type="hidden" name="checkstyle" value="<%=nowcheckstyle%>">
<input type="hidden" name="checkstatue" value="<%=nowcheckstatue%>">
<input type="hidden" name="keyword1" value="<%=nowkeyword1%>">
<table class='grid_table' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td width='100'>ID号</td>
	<td width='*'>用户名</td>
	<td width='150'>登录时间</td>
	<td width='150'>登陆IP</td>
	<td width='150'>事件</td>
	<td width='80'>删除</td>
</tr>
<%
set rs =server.createobject("ADODB.RecordSet")	
rs.open sql,conn,1,3
if not rs.eof then
	rs.pagesize=maxrecord
	rs.absolutepage=currentpage
	for currentrec=1 to rs.pagesize
    if rs.eof then
		exit for
    end if
	if thisClass = "grid_even" then thisClass = "grid_odd" else thisClass = "grid_even"
	%>
	<tr class=<%=thisClass%> align="center">
		<td class="grid_cell"><%=rs("id")%></td>
		<td class="grid_cell"><%=rs("username")%>（<%=shownamestr("userrealname","login","username",rs("username"),"")%>）</td>
		<td class="grid_cell"><font color=green><%=rs("logindate")%></font> </td>
		<td class="grid_cell"><%=rs("address")%></td>
		<td class="grid_cell"><%if rs("class")="登陆失败" then
			response.write "<font color='red'>"&rs("class")&"</font>"
		else
			response.write rs("class")
		end if%></td>
		<td class="grid_cell"><input type="checkbox" name="ID" value="<%=rs("id")%>" class="form_checkbox_normal"></td>
	</tr>
	<%
	Rs.movenext
	next
	
	if rs.recordcount>0 then 
	%>
	<tr class='grid_header'>
		<td colspan="8">
			<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
			<tr class="list_command">
				<td align="right">
				<%showdel="yes"%>
				<!--#include file="inc/page_bar.asp" --></td>
			</tr>
			</table>
		</td>
	</tr>
	<%end if
else
	response.write "<tr class='grid_span'><td colspan='8'><font color='red'>查无资料纪录！</font></td></tr>"
end if

rs.close
set rs = nothing
%>
</table>
</form>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->