<%nowneedtime1=request("needtime1")
nowneedtime2=request("needtime2")
nowneedtime3=request("needtime3")
nowneedtime4=request("needtime4")
nowstart_date=request("start_date") 
if nowstart_date="" then
  nowstart_date=date()-day(date()-1)
end if
nowend_date=request("end_date") 
if nowend_date="" then
  nowend_date=date()
end if

nowtender_class = request("tender_class")
nowcheckgrade = request("checkgrade")
nowcheckstyle = request("checkstyle")
nowcheckcondition = request("checkcondition")
nowcheckstatue = request("checkstatue")
nowkeyword1 = RemoveHTML(ReplaceQM(request("keyword1")))

gourl="needtime1="&nowneedtime1&"&needtime2="&nowneedtime2&"&needtime3="&nowneedtime3&"&needtime4="&nowneedtime4&"&start_date="&nowstart_date&"&end_date="&nowend_date&"&tender_class="&nowtender_class&"&checkgrade="&nowcheckgrade&"&checkstyle="&nowcheckstyle&"&checkcondition="&nowcheckcondition&"&checkstatue="&nowcheckstatue&"&keyword1="&nowkeyword1&""
%>
<table class='grid_search' cellpadding='0' cellspacing='1' border='0'>
<form name="form2">
<tr class='grid_header'>
	<td>时间选择</td>
	<td>查询条件</td>
	<td>关键字</td>
	<td>查询</td>
	<td>清除</td>
</tr>
<tr class="grid_odd" align="center">
	<td class="grid_cell">
	开始日期<input type="checkbox" name="needtime1" value="yes"<%if nowneedtime1="yes" then%> checked<%end if%>>
	截止日期<input type="checkbox" name="needtime2" value="yes"<%if nowneedtime2="yes" then%> checked<%end if%>>
	需求日期<input type="checkbox" name="needtime3" value="yes"<%if nowneedtime3="yes" then%> checked<%end if%>>
	发布日期<input type="checkbox" name="needtime4" value="yes"<%if nowneedtime4="yes" then%> checked<%end if%>>
	<input name="start_date" value="<%=nowstart_date%>" readonly onFocus="calendar()" style="width:100px">
	—
	<input name="end_date" value="<%=nowend_date%>" readonly onFocus="calendar()" style="width:100px">
	</td>
	<td class="grid_cell">
	<select name="tender_class" class='form_combo_normal' onChange="form2.submit();">
	<option value="">项目类型</option>
	<option value="材料"<%if nowtender_class="材料" then%> selected="selected"<%end if%>>材料</option>
	<option value="设备"<%if nowtender_class="设备" then%> selected="selected"<%end if%>>设备</option>
	<option value="综合"<%if nowtender_class="综合" then%> selected="selected"<%end if%>>综合</option>
	</select>
	<select name="checkgrade" class='form_combo_normal' onChange="form2.submit();">
	<%response.write "<option value=''>竞价级别</option>"
	response.write "<option value='0'"
	if nowcheckgrade="0" then response.write " selected"
	response.write ">所有等级</option>"
	set rsc=conn.execute("select * from companygrade where id<>0 order by companygradeorder")
	if not rsc.eof then
	do while not rsc.eof
	response.write "<option value="&rsc(0)
	if nowcheckgrade=trim(cstr(rsc(0))) then
		response.write " selected"
	end if
	response.write ">"&rsc(1)&"</option>"
	rsc.movenext
	loop
	end if
	rsc.close%>
	</select>
	<select name="checkcondition" class='form_combo_normal' onChange="form2.submit();">
	<option value="">查询条件</option>
	<option value="1"<%if nowcheckcondition="1" then%> selected="selected"<%end if%>>按项目名称</option>
	<option value="2"<%if nowcheckcondition="2" then%> selected="selected"<%end if%>>按竞价编号</option>
	<option value="3"<%if nowcheckcondition="3" then%> selected="selected"<%end if%>>按采购依据</option>
	<option value="4"<%if nowcheckcondition="4" then%> selected="selected"<%end if%>>按详情说明</option>
	<option value="5"<%if nowcheckcondition="5" then%> selected="selected"<%end if%>>按物资名称</option>
	<option value="6"<%if nowcheckcondition="6" then%> selected="selected"<%end if%>>按发布人</option>
	</select>
	<select name="checkstyle" class='form_combo_normal' onChange="form2.submit();">
	<option value="mohu"<%if nowcheckstyle="mohu" then%> selected="selected"<%end if%>>模糊</option>
	<option value="jing"<%if nowcheckstyle="jing" then%> selected="selected"<%end if%>>精确</option>
	</select></td>
	<td class="grid_cell"><input type='text' class='form_textbox_normal' name='keyword1' value='<%=nowkeyword1%>' size='15' ></td>
	<td class="grid_cell"><input type='submit' value='查询' class='form_button'></td>
	<td class="grid_cell"><input type='button' value='清除' class='form_button' onClick="window.location.href='?'"></td>
</tr>
</form>
</table>