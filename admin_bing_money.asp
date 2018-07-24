<!--#include file="Conn.asp" -->
<!--#include file="Const.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
if session("iflogin")>0 then
call WhereTable("yast_partitioner.png","竞价总额饼状分析图")

nowstartyear=request("startyear")
if nowstartyear="" then nowstartyear=year(date())
nowstartmonth=request("startmonth")
if nowstartmonth="" then nowstartmonth=1
nowendyear=request("endyear")
if nowendyear="" then nowendyear=year(date())
nowendmonth=request("endmonth")
if nowendmonth="" then nowendmonth=month(date())

allno=datediff("m",nowstartyear&"-"&nowstartmonth&"-1",nowendyear&"-"&nowendmonth&"-1")+1
startyear=nowstartyear
startmonth=nowstartmonth

'参数含义(数组，横坐标，纵坐标，图表的宽度，图表的高度,图表标题,单位)
function table2(stat_array,table_left,table_top,all_width,all_height,table_title,unit)
 dim bg_color(3),pie(3)
 bg_color(1)="#ff1919"
 bg_color(2)="#ffff19"
 bg_color(3)="#1919ff"
 num =ubound(stat_array,1)
 allvalues=0
 for i=1 to num
 allvalues = allvalues+stat_array(i,1)
 next
 if allvalues=0 then response.end
 k=0
 for i=1 to num-1
 pie(i)=formatnumber(stat_array(i,1)/allvalues,4,-1)
 k=k+pie(i)
 next
 pie(num)=formatnumber((1-k),4,-1)
 response.Write "<v:shapetype id='Cake_3D' coordsize='21600,21600' o:spt='95' adj='11796480,5400' path='al10800,10800@0@0@2@14,10800,10800,10800,10800@3@15xe'></v:shapetype>"
 response.Write "<v:shapetype id='3dtxt' coordsize='21600,21600' o:spt='136' adj='10800' path='m@7,l@8,m@5,21600l@6,21600e'> "
 response.Write " <v:path textpathok='t' o:connecttype='custom' o:connectlocs='@9,0;@10,10800;@11,21600;@12,10800' o:connectangles='270,180,90,0'/>"
 response.Write " <v:textpath on='t' fitshape='t'/>"
 response.Write " <o:lock v:ext='edit' text='t' shapetype='t'/>"
 response.Write "</v:shapetype>"
 'response.Write "<v:rect id='background' style='position:absolute;left:"&table_left&"px;top:"&table_top&"px;WIDTH:"&all_width&"px;HEIGHT:"&all_height&"px;' fillcolor='#EFEFEF' strokecolor='gray'>"
 'response.Write " <v:shadow on='t' type='single' color='silver' offset='4pt,4pt'/>"
 'response.Write "</v:rect>"
 response.Write "<v:group ID='table' style='position:absolute;left:"&table_left&"px;top:"&table_top&"px;WIDTH:"&all_width&"px;HEIGHT:"&all_height&"px;' coordsize = '21000,11500'>"
' response.Write " <v:Rect style='position:relative;left:500;top:200;width:20000;height:800'filled='false' stroked='false'>"
' response.Write " <v:TextBox inset='0pt,0pt,0pt,0pt'>"
' response.Write " <table width='100%' border='0' align='center' cellspacing='0'>"
' response.Write " <tr>"
' response.Write " <td align='center' valign='middle'><div style='font-size:15pt; font-family:黑体;'><B>"&table_title&"</B></div></td>"
' response.Write " </tr>"
' response.Write " </table>"
' response.Write " </v:TextBox>"
' response.Write " </v:Rect> "
 response.Write " <v:rect id='back' style='position:relative;left:500;top:1000;width:20000; height:10000;' onmouseover='movereset(1)' onmouseout='movereset(0)' fillcolor='#9cf' strokecolor='#888888'>"
 response.Write " <v:fill rotate='t' angle='-45' focus='100%' type='gradient'/>"
 response.Write " </v:rect>"
 response.Write " <v:rect id='back' style='position:relative;left:15000;top:1400;width:5000; height:"&((num+1)*9000/11+200)&";' fillcolor='#9cf' stroked='t' strokecolor='#0099ff'>"
 response.Write " <v:fill rotate='t' angle='-175' focus='100%' type='gradient'/>"
 response.Write " <v:shadow on='t' type='single' color='silver' offset='3pt,3pt'/>"
 response.Write " </v:rect>"
 response.Write " <v:Rect style='position:relative;left:15500;top:1500;width:4000;height:700' fillcolor='#000000' stroked='f' strokecolor='#000000'>"
 response.Write " <v:TextBox inset='8pt,4pt,3pt,3pt' style='font-size:11pt;'><div align='left'><font color='#ffffff'><B>总数:"&allvalues&unit&"</B> </font></div></v:TextBox>"
 response.Write " </v:Rect> "
 for i=1 to num
 response.Write " <v:Rect id='rec"&i&"' style='position:relative;left:15400;top:"&i*9000/11+1450&";width:4300;height:800;display:none' fillcolor='#efefef' strokecolor='"&bg_color(i)&"'>"
 response.Write " <v:fill opacity='.6' color2='fill darken(118)' o:opacity2='.6' rotate='t' method='linear sigma' focus='100%' type='gradient'/>"
 response.Write " </v:Rect>"
 response.Write " <v:Rect style='position:relative;left:15500;top:"&i*9000/11+1500&";width:600;height:700' fillcolor='"&bg_color(i)&"' stroked='f'/>"
 response.Write " <v:Rect style='position:relative;left:16300;top:"&i*9000/11+1500&";width:3400;height:700' filled='f' stroked='f'>"
 response.Write " <v:TextBox inset='0pt,5pt,0pt,0pt' style='font-size:9pt;'><div align='left'>"&stat_array(i,2)&":"&stat_array(i,1)&unit&"</div></v:TextBox>"
 response.Write " </v:Rect> "
 next
 response.Write "</v:group>"
 k1=180
 k4=10
 for i=1 to num
'response.write "<a href=aaa.asp>"
 k2=360*pie(i)/2
 k3=k1+k2
 if k3>=360 then
 k3=k3-360
 end if
 kkk=(-11796480*pie(i)+5898240)
 k5=3.1414926*2*(180-(k3-180))/360
 R=all_height/2
 txt_x = table_left+all_height/8-30+R+R*sin(k5)*0.7
 txt_y = table_top+all_height/14-39+R+R*cos(k5)*0.7*0.5
 titlestr = "名称："&stat_array(i,2)&" 数值："&stat_array(i,1)&unit&" 所占比例："&pie(i)*100&"%"
 response.Write " <div style='cursor:hand;'>"
 response.Write " <v:shape id='cake"&i&"' type='#Cake_3D' title='"&titlestr&"'"
 response.Write " style='position:absolute;left:"&table_left+all_height/8&"px;top:"&table_top+all_height/14&"px;WIDTH:"&all_height&"px;HEIGHT:"&all_height&"px;rotation:"&k3&";z-index:"&k4&"'"
 response.Write " adj='"&kkk&",0' fillcolor='"&bg_color(i)&"' onmouseover='moveup(cake"&i&","&(table_top+all_height/14)&",txt"&i&",rec"&i&")'; onmouseout='movedown(cake"&i&","&(table_top+all_height/14)&",txt"&i&",rec"&i&");'>"
 response.Write " <v:fill opacity='60293f' color2='fill lighten(120)' o:opacity2='60293f' rotate='t' angle='-135' method='linear sigma' focus='100%' type='gradient'/>"
 response.Write " <o:extrusion v:ext='view' on='t'backdepth='25' rotationangle='60' viewpoint='0,0'viewpointorigin='0,0' skewamt='0' lightposition='-50000,-50000' lightposition2='50000'/>"
 response.Write " </v:shape>"
 response.Write " <v:shape id='txt"&i&"' type='#3dtxt' style='position:absolute;left:"&txt_x&"px;top:"&txt_y&"px;z-index:20;display:none;width:50; height:18;' fillcolor='#ffffff'"
 response.Write " onmouseover='ontxt(cake"&i&","&(table_top+all_height/14)&",txt"&i&",rec"&i&")'>"
 response.Write " <v:fill opacity='60293f' color2='fill lighten(120)' o:opacity2='60293f' rotate='t' angle='-135' method='linear sigma' focus='100%' type='gradient'/>"
 response.Write " <v:textpath style='font-family:'宋体';v-text-kern:t' trim='t' fitpath='t' string='"&pie(i)*100&"%'/>"
 response.Write " <o:extrusion v:ext='view' backdepth='8pt' on='t' lightposition='0,0' lightposition2='0,0'/>"
 response.Write " </v:shape>"
 response.Write " </div>"
 k1=k1+k2*2
 if k1>=360 then
 k1=k1-360
 end if
 if k1>180 then
 k4=k4+1
 else
 k4=k4-1
 end if
'response.write "</a>"
 next
end function
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<link href='css/workspace.css' rel='stylesheet' type='text/css'>
<title></title>
<STYLE>
v\:* { Behavior: url(#default#VML) }
o\:* { behavior: url(#default#VML) }
</STYLE>
</head>
<body>
<script language="JavaScript" src="js/Function.js"></script>
<SCRIPT language="JavaScript" src="inc/day.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
<!--
onit=true
num=0
function moveup(iteam,top,txt,rec){
temp=eval(iteam)
tempat=eval(top)
temptxt=eval(txt)
temprec=eval(rec)
at=parseInt(temp.style.top)
temprec.style.display = "";
if (num>27){
temptxt.style.display = "";
}
if(at>(tempat-28)&&onit){
num++
temp.style.top=at-1
Stop=setTimeout("moveup(temp,tempat,temptxt,temprec)",10)
}else{
return
}
}
function movedown(iteam,top,txt,rec){
temp=eval(iteam)
temptxt=eval(txt)
temprec=eval(rec)
clearTimeout(Stop)
temp.style.top=top
num=0
temptxt.style.display = "none";
temprec.style.display = "none";
}
function ontxt(iteam,top,txt,rec){
temp = eval(iteam);
temptxt = eval(txt);
temprec = eval(rec)
if (onit){
temp.style.top = top-28;
temptxt.style.display = "";
temprec.style.display = "";
}
}
function movereset(over){
if (over==1){
onit=false
}else{
onit=true
}
}
-->
</script>
<script type="text/javascript"><!--
google_ad_client = "pub-3971858006444213";
google_ad_slot = "7286086700";
google_ad_width = 468;
google_ad_height = 15;
//--></script>
<%
dim total(3,2)

if isSqldatabase=0 then
	sqltime=" and gettime between #"&getmonthstartdate(nowstartyear,nowstartmonth)&"# and #"&getmonthenddate(nowendyear,nowendmonth)+1&"#"
else
	sqltime=" and gettime between '"&getmonthstartdate(nowstartyear,nowstartmonth)&"' and '"&getmonthenddate(nowendyear,nowendmonth)+1&"'"
end if

total(1,1)=getshuliang1("spotprice","competitive","(statue=1 or statue=2) and ifzu=1 and djh in (select djh from tender where tenderclass='材料')"&sqltime)
total(2,1)=getshuliang1("spotprice","competitive","(statue=1 or statue=2) and ifzu=1 and djh in (select djh from tender where tenderclass='设备')"&sqltime)
total(3,1)=getshuliang1("spotprice","competitive","(statue=1 or statue=2) and ifzu=1 and djh in (select djh from tender where tenderclass='综合')"&sqltime)
total(1,2)="材料"
total(2,2)="设备"
total(3,2)="综合"

allzhaobiao2=total(1,1)
allzhaobiao3=total(2,1)
allzhaobiao4=total(3,1)
allzhaobiao1=allzhaobiao2+0+allzhaobiao3+0+allzhaobiao4%>

<table class='grid_search' cellpadding='0' cellspacing='1' border='0'>
<form name="form2">
<tr class='grid_header'>
	<td>查询条件</td>
	<td>查询</td>
	<td>清除</td>
	<td width="*"></td>
</tr>
<tr class="grid_odd" align="center">
	<td class="grid_cell">
	<select name="startyear" class='form_combo_normal' onChange="form2.submit();">
	<option value="">选择起始年</option>
	<%for i=year(date())-2 to year(date())
		response.write "<option value="&i
		if nowstartyear-i=0 then response.write " selected"
		response.write ">"&i&"年</option>"
	next%>
	</select>
	<select name="startmonth" class='form_combo_normal' onChange="form2.submit();">
	<option value="">选择起始月</option>
	<%for i=1 to 12
		response.write "<option value="&i
		if nowstartmonth-i=0 then response.write " selected"
		response.write ">"&i&"月</option>"
	next%>
	</select>
	
	至
	
	<select name="endyear" class='form_combo_normal' onChange="form2.submit();">
	<option value="">选择终止年</option>
	<%for i=year(date())-5 to year(date())
		response.write "<option value="&i
		if nowendyear-i=0 then response.write " selected"
		response.write ">"&i&"年</option>"
	next%>
	</select>
	<select name="endmonth" class='form_combo_normal' onChange="form2.submit();">
	<option value="">选择终止月</option>
	<%for i=1 to 12
		response.write "<option value="&i
		if nowendmonth-i=0 then response.write " selected"
		response.write ">"&i&"月</option>"
	next%>
	</select></td>
	<td class="grid_cell"><input type='submit' value='查询' class='form_button'></td>
	<td class="grid_cell"><input type='button' value='清除' class='form_button' onClick="window.location.href='?'"></td>
	<td class="grid_cell"></td>
</tr>
<tr class='grid_header'>
	<td colspan="4"><strong>竞价项目：<%=allzhaobiao1%>&nbsp;&nbsp;&nbsp;&nbsp;材料类：<%=allzhaobiao2%>&nbsp;&nbsp;&nbsp;&nbsp;设备类：<%=allzhaobiao3%>&nbsp;&nbsp;&nbsp;&nbsp;综合类：<%=allzhaobiao4%></strong></td>
</tr>
</form>
</table>
<%
call table2(total,120,120,800,420,"竞价数量三维饼状图","元")
'参数含义(数组，横坐标，纵坐标，图表的宽度，图表的高度,图表标题,单位)%>
<%else
response.redirect "erro.asp"
response.end
end if%>
</body>
</html>