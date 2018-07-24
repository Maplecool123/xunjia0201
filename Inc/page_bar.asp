<%response.write rs.recordcount&"&nbsp;条信息&nbsp; 共&nbsp;"&rs.pagecount&"&nbsp;页&nbsp;"
nowstart=currentpage-3
if nowstart<1 then
    nowstart=1
end if
nowend=currentpage+3
if nowend>rs.pagecount then
    nowend=rs.pagecount
end if
response.write "&nbsp;<a href='?page=1&"&gourl&"'><img src='images/page_first.png' border=0></a>&nbsp;"
if currentpage-1>0 then
	response.write "&nbsp;<a href='?page="&currentpage-1&"&"&gourl&"'><img src='images/page_prev.png' border=0></a>&nbsp;"
end if
for ipage=nowstart to nowend
	if cstr(ipage)=cstr(currentpage) then
		response.write "&nbsp;<span style='font-weight:bold;color:#5858E6'>" & ipage &"</span>&nbsp;"
    else
		response.write "&nbsp;[&nbsp;<a href='?page="&ipage&"&"&gourl&"' class='page'>" & ipage &"</a>&nbsp;]&nbsp;"
    end if
next
if currentpage-rs.pagecount<0 then
	response.write "&nbsp;<a href='?page="&currentpage+1&"&"&gourl&"'><img src='images/page_next.png' border=0></a>&nbsp;"
end if
response.write "&nbsp;<a href='?page="&rs.pagecount&"&"&gourl&"'><img src='images/page_last.png' border=0></a>&nbsp;"

response.write "<select name='page' onChange=""form1.action='?';form1.submit();"" class='form_combo_normal'>"

for x=1 to rs.pagecount
	response.write "<option value="&x
	if currentpage=trim(cstr(x)) then
		response.write " selected"
	end if
	response.write ">"&x&"</option>"
next

response.write "</select>"

if showdel="yes" then
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chkall' id='chkall' value='select' onClick='CheckAll(this.form)' class='form_checkbox_normal'> 全选&nbsp;"
	response.write "<input type='submit' value='删 除' onClick=""return confirm('确定要删除记录吗？')"" class='form_submit'>&nbsp;&nbsp;"
end if%>