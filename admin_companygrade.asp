<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
'全局变量
Dim CurrentPage,sql,i,rs
'全局变量

call WhereTable("todo.png","供应商资质级别管理")

dim nowcheckstyle,nowkeyword1
currentpage=request("page")

'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

nowcheckstyle = request("checkstyle")
nowkeyword1 = RemoveHTML(ReplaceQM(request("keyword1")))

gourl="checkstyle="&nowcheckstyle&"&keyword1="&nowkeyword1&""

if request("action")="delall" then
	id=replace(request("id")," ","")

	goon=0
	if replace(id,",","")<>"" then goon=1

	if goon=0 then
		call HintAndBack("请至少选中一条记录！",1)
	end if
	
	id=split(id,",")
	for i=0 to UBound(id)
		conn.execute("delete from [companygrade] where id="&id(i)&"")
	next
	
	call HintAndTurn("删除成功！","?"&gourl&"&page="&currentpage&"")
end if
%>
<table class='grid_search' cellpadding='0' cellspacing='1' border='0'>
<form name="form2">
<tr class='grid_header'>
	<td width='120'>查询条件</td>
	<td width='140'>关键字</td>
	<td width='70'>查询</td>
	<td width='70'>清除</td>
	<td width='70'>新增级别</td>
	<td width='*'></td>
</tr>
<tr class="grid_odd" align="center">
	<td class="grid_cell"><select name="checkstyle" class='form_combo_normal' onChange="form2.submit();">
	<option value="mohu"<%if nowcheckstyle="mohu" then%> selected="selected"<%end if%>>模糊</option>
	<option value="jing"<%if nowcheckstyle="jing" then%> selected="selected"<%end if%>>精确</option>
	</select></td>
	<td class="grid_cell"><input type='text' class='form_textbox_normal' name='keyword1' value='<%=nowkeyword1%>' size='15' ></td>
	<td class="grid_cell"><input type='submit' value='查询' class='form_button'></td>
	<td class="grid_cell"><input type='button' value='清除' class='form_button' onClick="window.location.href='?'"></td>
	<td class="grid_cell"><input type='button' value='新增' onClick="location.href='admin_companygrade_add.asp'" class='form_submit'></td>
	<td class="grid_cell" align="left"></td>
</tr>
</form>
</table>
<%
dim thisClass
thisClass = "grid_even"

Set Rs = server.createobject("adodb.recordset")
sql="select * from companygrade where id<>0"
if nowcheckstyle="mohu" then
	if nowkeyword1<>"" then
		sql=sql&" and companygradename like '%"&nowkeyword1&"%'"
	end if
elseif nowcheckstyle="jing" then
	if nowkeyword1<>"" then
		sql=sql&" and companygradename = '"&nowkeyword1&"'"
	end if
end if
sql=sql&" order by companygradeorder"%>
<form name='form1'>
<input type="hidden" name="action" value="delall">
<input type="hidden" name="checkstyle" value="<%=nowcheckstyle%>">
<input type="hidden" name="keyword1" value="<%=nowkeyword1%>">
<table class='grid_table' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td width='100'>ID号</td>
	<td width='*'>供应商级别名称</td>
	<td width='250'>级别高低（级别越高者，数值越大）</td>
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
		<td class="grid_cell"><%=rs("companygradename")%></td>
		<td class="grid_cell"><%=rs("companygradeorder")%></td>
		<td class="grid_cell"><input type='button' value='修改' onClick="location.href='admin_companygrade_edit.asp?id=<%=rs("id")%>&<%=gourl%>&page=<%=currentpage%>'" class='form_submit'></td>
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