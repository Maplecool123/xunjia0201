<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
'全局变量
Dim CurrentPage,sql,i,rs
'全局变量

call WhereTable("administration.png","管理员管理")

dim nowgrade,nowcheckclass,nowcheckstyle,nowkeyword
currentpage=request("page")
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

nowgrade = request("grade")
nowcheckclass = request("checkclass")
nowcheckstyle = request("checkstyle")
nowkeyword1 = RemoveHTML(ReplaceQM(request("keyword1")))

gourl="grade="&nowgrade&"&checkclass="&nowcheckclass&"&checkstyle ="&nowcheckstyle &"&keyword1="&nowkeyword1&""

if request("action")="delall" then
	id=replace(request("id")," ","")

	goon=0
	if replace(id,",","")<>"" then goon=1

	if goon=0 then
		call HintAndBack("请至少选中一条记录！",1)
	end if
	
	id=split(id,",")
	for i=0 to UBound(id)
		conn.execute("delete from [login] where id="&id(i)&"")
	next
	
	
	call HintAndTurn("删除成功！","?"&gourl&"&page="&currentpage&"")
end if
%>
<table class='grid_search' cellpadding='0' cellspacing='1' border='0'>
<form name="form2">
<tr class='grid_header'>
	<td width='300'>查询条件</td>
	<td width='140'>关键字</td>
	<td width='70'>查询</td>
	<td width='70'>清除</td>
	<td width='90'>新增管理员</td>
	<td width='*'></td>
</tr>
<tr class="grid_odd" align="center">
	<td class="grid_cell">
	<select name="grade" class='form_combo_normal' onChange="form2.submit();">
	<option value="">管理员级别</option>
	<option value="1"<%if nowgrade="1" then%> selected="selected"<%end if%>>计划员</option>
	<option value="2"<%if nowgrade="2" then%> selected="selected"<%end if%>>注册审核员</option>
	<option value="3"<%if nowgrade="3" then%> selected="selected"<%end if%>>纪管审查员</option>
	<option value="99"<%if nowgrade="99" then%> selected="selected"<%end if%>>超级管理员</option>
	</select>
	<select name="checkclass" class='form_combo_normal' onChange="form2.submit();">
	<option value="1"<%if nowcheckclass="1" then%> selected="selected"<%end if%>>按用户名</option>
	<option value="2"<%if nowcheckclass="2" then%> selected="selected"<%end if%>>按真实姓名</option>
	<option value="3"<%if nowcheckclass="3" then%> selected="selected"<%end if%>>按联系电话</option>
	</select>
	<select name="checkstyle" class='form_combo_normal' onChange="form2.submit();">
	<option value="mohu"<%if nowcheckstyle="mohu" then%> selected="selected"<%end if%>>模糊</option>
	<option value="jing"<%if nowcheckstyle="jing" then%> selected="selected"<%end if%>>精确</option>
	</select></td>
	<td class="grid_cell"><input type='text' class='form_textbox_normal' name='keyword1' value='<%=nowkeyword1%>' size='15' ></td>
	<td class="grid_cell"><input type='submit' value='查询' class='form_button'></td>
	<td class="grid_cell"><input type='button' value='清除' class='form_button' onClick="window.location.href='?'"></td>
	<td class="grid_cell"><input type='button' value='新增' onClick="location.href='admin_login_add.asp'" class='form_submit'></td>
	<td class="grid_cell" align="left"></td>
</tr>
</form>
</table>
<%
dim thisClass
thisClass = "grid_even"

sql="select * from login where id<>0"
if nowgrade<>"" then
	sql=sql&" and usergrade="&nowgrade&""
end if
if nowkeyword1<>"" then
	if nowcheckstyle="mohu" then
		select case nowcheckclass
		case "1"
		sql=sql&" and username like '%"&nowkeyword1&"%'"
		case "2"
		sql=sql&" and userrealname like '%"&nowkeyword1&"%'"
		case "3"
		sql=sql&" and tel like '%"&nowkeyword1&"%'"
		end select
	elseif nowcheckstyle="jing" then
		select case nowcheckclass
		case "1"
		sql=sql&" and username = '"&nowkeyword1&"'"
		case "2"
		sql=sql&" and userrealname = '"&nowkeyword1&"'"
		case "3"
		sql=sql&" and tel = '"&nowkeyword1&"'"
		end select
	end if
end if
sql=sql&" order by usergrade"%>
<form name='form1'>
<input type="hidden" name="action" value="delall">
<input type="hidden" name="grade" value="<%=nowgrade%>">
<input type="hidden" name="checkclass" value="<%=nowcheckclass%>">
<input type="hidden" name="checkstyle" value="<%=nowcheckstyle%>">
<input type="hidden" name="keyword1" value="<%=nowkeyword1%>">
<table class='grid_table' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td width='100'>ID号</td>
	<td width='*'>用户名</td>
	<td width='150'>真实姓名</td>
	<td width='150'>联系电话</td>
	<td width='150'>级别</td>
	<td width='80'>修改</td>
	<td width='80'>删除</td>
</tr>
<%
set rs =server.createobject("ADODB.RecordSet")	
rs.open sql,conn,1,3
if not rs.eof then
	rs.pagesize=maxrecord1
	rs.absolutepage=currentpage
	for currentrec=1 to rs.pagesize
    if rs.eof then
		exit for
    end if
	if thisClass = "grid_even" then thisClass = "grid_odd" else thisClass = "grid_even"
	%>
	<tr class=<%=thisClass%> align="center">
		<td class="grid_cell"><%=rs("id")%></td>
		<td class="grid_cell"><%=rs("username")%></td>
		<td class="grid_cell"><font color=green><%=rs("userrealname")%></font> </td>
		<td class="grid_cell"><%=rs("tel")%></td>
		<td class="grid_cell"><%=showadmingrade(rs("usergrade"))%></td>
		<td class="grid_cell"><input type='button' value='修改' onClick="location.href='admin_login_edit.asp?id=<%=rs("id")%>&<%=gourl%>&page=<%=currentpage%>'" class='form_submit'></td>
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