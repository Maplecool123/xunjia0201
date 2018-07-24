
用纯ASP实现完美的WEB柱状图
[ 2007-11-30 14:15:30 | 编辑&整理: Admin ]
字号: 大 | 中 | 小
<%
dim total(7,2)
total(1,1)=200
total(2,1)=800
total(3,1)=1000
total(4,1)=600
total(5,1)=1222
total(6,1)=3213
total(7,1)=8

total(1,2)="中国经营报"
total(2,2)="招聘网"
total(3,2)="51Job"
total(4,2)="新民晚报"
total(5,2)="新闻晚报"
total(6,2)="南方周末"
total(7,2)="羊城晚报"

total_no=7
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<!--[if !mso]>
<style>
v\:* { behavior: url(#default#VML) }
o\:* { behavior: url(#default#VML) }
.shape { behavior: url(#default#VML) }
</style>
<![endif]-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link rel="stylesheet" href="List.css"></head>
<body topmargin=5 leftmargin=0 scroll=no>
<%call table1(total,20,15,470,200)%>
</body>
</html>

以上是调用函数的例子，下面是所调用的函数
<%
function table1(total,thickness,table_space,all_width,all_height)
'参数含义(传递的数组，柱子的厚度，柱子的间隔，图表的宽度，图表的高度)
'纯ASP代码生成图表函数1——柱状图
'作者：龚鸣(Passwordgm) QQ:25968152 MSN:passwordgm@sina.com Email:passwordgm@sina.com
'本人非常愿意和ASP,VML,FLASH的爱好者在HTTP://topclouds.126.com进行交流和探讨
'版本1.0 最后修改日期 2003-7-10
'非常感谢您使用这个函数，请您使用和转载时保留版权信息，这是对作者工作的最好的尊重。
dim tb_color(7,2)

tb_color(1,1)="#d1ffd1"
tb_color(2,1)="#ffbbbb"
tb_color(3,1)="#ffe3bb"
tb_color(4,1)="#cff4f3"
tb_color(5,1)="#d9d9e5"
tb_color(6,1)="#ffc7ab"
tb_color(7,1)="#ecffb7"

tb_color(1,2)="#00ff00"
tb_color(2,2)="#ff0000"
tb_color(3,2)="#ff9900"
tb_color(4,2)="#33cccc"
tb_color(5,2)="#666699"
tb_color(6,2)="#993300"
tb_color(7,2)="#99cc00"

response.write "<table border=0 cellpadding=0 cellspacing=0 style=border-collapse: collapse bordercolor=#111111 width="&all_width&" height="&all_height&">"
response.write "<tr><td width=100% height=* valign=middle><table border=0 cellpadding=0 cellspacing=0 style=border-collapse: collapse bordercolor='#111111' width='100%' height='100%'>"
response.write "<tr align='center'><td width='35' height='100%' valign='bottom'>"
response.write "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' height='100%'>"

temp1=0
for i=1 to total_no
if temp1<total(i,1) then temp1=total(i,1)
next
temp1=int(temp1)
if temp1>9 then
temp2=mid(cstr(temp1),2,1)
if temp2>4 then
temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+1)*10^(len(cstr(temp1))-1)
else
temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+0.5)*10^(len(cstr(temp1))-1)
end if
else
if temp1>4 then temp3=10 else temp3=5
end if
if total_no>0 then
for i=temp3 to 1 step -temp3/5
response.write "<tr style='font-size:1px;height:1px'><td></td><td bgcolor='#111111' width='20%'></td></tr>"
response.write "<tr align=right valign='top'><td colspan='2'>"&i&"</td></tr>"
next
response.write "</table>"
response.write "</td><td style='font-size:1px;height:1px' bgcolor='#111111'>1</td><td width='"&(all_width-30)&"' height='"&(all_height-30)&"' valign='bottom' align='left'>"
response.write "<!--[if gte vml 1]>"

z=9
width=30
total_width=280
width=(total_width-total_no*z*2)/(total_no)
m=0
if width>30 then width=30
m=m+1
for i=1 to total_no
response.write "<v:rect id='_x0000_s1025' alt='' style='position:relative;left:"
response.write table_space/2+table_space*(i-1)
response.write "pt;top:2px;width:"&width&"pt;height:"&(all_height/1.41)*total(i,1)/temp3&"pt;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' type='gradient'/>"
response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
response.write "<v:textbox inset='0,0,0,0'>"
response.write "<table cellspacing=0 cellpadding=0 width='100%' height='100%'>"
response.write "<tr><td align='center'"
if (all_height/1.41)*total(i,1)/temp3<8 then response.write " style='font-size:1px;'"
response.write ">"&total(i,1)&"</td></tr>"
response.write "</table></v:textbox></v:rect>"
next

response.write "<![endif]--></td></tr>"
response.write "<tr align='center'><td></td><td style='font-size:1px;height:1px' bgcolor='#111111'></td><td style='font-size:1px;height:1px' bgcolor='#111111'></td></td>"
response.write "<tr align='center'><td></td><td></td><td width='' height='*' valign='middle'>"
response.write "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' height='30'>"
response.write "<tr align=center valign='center'>"

for i=1 to total_no
response.write "<td width='"&(100/total_no)&"%'>"&total(i,2)&"</td>"
next
else
response.write ""
end if
response.write "</tr></table></td></tr></table></td></tr></table>"
end function
%>