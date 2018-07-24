<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
call WhereTable("todo.png","系统初始化")
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="4"><strong>系统初始化确认</strong></td>
</tr>
<form action="New_submit.asp" method="post">
<tr>
	<td align="center" class="list_required">
		<strong>警   告！</strong><br><br>
		<strong>初始化一般在系统第一次安装时进行！</strong><br>
		<strong>在系统正常运行期间初始化将可能丢失数据。</strong><br><br>
系统将被初始化，会删除数据库中的所有信息，但仍会保留数据库结构，确认要对系统做<font color=red>完全初始化</font>操作，请按“确定”，否则请按“取消”！
	</td>
</tr>
<tr>
	<td width="100%" align="center" class="list_command"> 
		<input type="hidden" name="action" value='del1'>
		<input type="submit" class='form_submit' value="确定" name="alert_button">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" class='form_submit' value="取消" name="cancel" onClick="javascript:history.go(-1)">
	</td>
</tr>
</form>
</table>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td colspan="4">&nbsp;</td>
</tr>
</table>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<form action="New_submit.asp" method="post">
<tr>
	<td class="grid_header" colspan="4"><strong>你也可以选择<font color=red>单独初始化</font>系统的部分</strong></td>
</tr>
<tr align="center">
	<td width="25%" class="list_required"><input type="checkbox" class='form_textbox_normal' value="gonggao" name="selectdel">系统公告</td>
	<td width="25%" class="list_required"><input type="checkbox" class='form_textbox_normal' value="yijian" name="selectdel">意见建议</td>
	<td width="25%" class="list_required"><input type="checkbox" class='form_textbox_normal' value="rizi" name="selectdel">系统日志</td>
	<td width="25%" class="list_required"><input type="checkbox" class='form_textbox_normal' value="companygrade" name="selectdel">资质级别</td>
</tr>
<tr align="center">
	<td class="list_required"><input type="checkbox" class='form_textbox_normal' value="companytype" name="selectdel">公司类型</td>
	<td class="list_required"><input type="checkbox" class='form_textbox_normal' value="company" name="selectdel">供 应 商</td>
	<td class="list_required"><input type="checkbox" class='form_textbox_normal' value="tender" name="selectdel">竞价信息</td>
	<td class="list_required"><input type="checkbox" class='form_textbox_normal' value="competitive" name="selectdel">竞价数据</td>
</tr>
<tr>
<td colspan="4"  height="25" align="center" class="list_command">
	<input type="hidden" name="action" value='del2'>
	<input type="checkbox" name="chkall" class='form_textbox_normal' value="on" onclick="CheckAll(this.form)">全选&nbsp;&nbsp;
	<input type="submit" name="action1" class='form_submit' onclick="{if(confirm('初始化选定的项目吗？')){this.document.check;return true;}return false;}" value="初始化">
</td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->