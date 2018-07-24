<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-99=0 then
call WhereTable("gnome-ccdesktop.png","导入竞价信息第二步：导入物资信息")%>

<script language="JavaScript" type="text/JavaScript">
function chk()
{
  if (document.form1.ExName.value=="")
  {
    alert("Excel文件名称不能为空！");
    document.form1.ExName.focus();
    return false;
  }
      if (document.form1.ExTname.value=="")
  {
    alert("Excel数据表文件名称不能为空！");
    document.form1.ExTname.focus();
    return false;
  }
}
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td colspan="4">导入说明</td>
</tr>
<tr class="grid_odd">
	<td colspan="4">
	<b>导入数据注意事项</b><p>
	1:您首先应当完成系统“参数设置”中各个参数的基础配置工作。<p>
	2:如果您的系统数据库里已经有部分数据，建议您先<a href="data_s.asp?action=BackupData"><font color="#FF0000">备份数据库</font></a>。<p>
	3:请确保你清楚Excel文件内容字段与导入数据库的字段相同。<p>
	4:请确保你清楚Excel文件的表名正确 如 Sheet1（默认）。<p>
	5:必须严格按照模板的格式来制作待导入的Excel表格，下载模板：<a href="mother.xls" target="_blank"><font color="#FF0000">下载</font></a><p>
	6:如果有不明白，或者需要测试导入功能，请下载范例Excel模板，了解怎样正确填写待导入的Excel表：<a href="example.xls" target="_blank"><font color="#FF0000">下载</font></a><p>
	<b>导入数据操作步骤</b><p>
	1:点击“浏览”选中待导入的Excel表，然后点击“上传”。<p>
	2:显示"Excel file uploads successful!"后，点击“导入数据”。选中“更新数据库里已有信息”可以把数据库里已有的数据更新成导入的Excel表里的信息。<p>
	</td>
</tr>
		
<form method="post" action="upfilepro.asp" name="form2">
<tr>
	<td width="15%" height="29" align="center" class="list_header_required"><strong>第一步</strong></td>
	<td align="left" width="30%" class="list_required">&nbsp;上传需要导入的Excel报表</td>
	<td colspan="2" align="left" class="list_required">
	<IFRAME src="upfilepro.asp" frameBorder=0 width="100%" height="29" scrolling=no align="top"></IFRAME>
	</td>
</tr>
</form>
<form name="form1" method="post" action="admin_tender_add_excelsave.asp" onSubmit="return chk(this)">
<input type="hidden" name="djh" id="djh" value="<%=request("djh")%>">
<tr>
	<td rowspan="2" align="center" class="list_header_required"><strong>第二步</strong></td>
	<td height="29" class="list_required">&nbsp;Excel名称</td>
	<td width="30%" class="list_required">&nbsp;需要导入的Excel数据表名 (如:Sheet1)</td>
	<td rowspan="2" class="list_command" >&nbsp;<input type="submit" value="导入数据" class='form_submit'></td>
</tr>
<tr>
	<td height="29" class="list_required">&nbsp;<input type="text" name="ExName" id="ExName" value="" style="width:200px;" class='form_textbox_normal'></td>
	<td class="list_required">&nbsp;<input type="text" name="ExTname" id="ExTname" value="Sheet1" style="width:200px;" class='form_textbox_normal'></td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->