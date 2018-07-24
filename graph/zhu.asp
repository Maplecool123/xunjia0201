<%
dim total(12,2)

total(1,1)=56
total(2,1)=7
total(3,1)=-7
total(4,1)=0
total(5,1)=16
total(6,1)=10
total(7,1)=14
total(8,1)=7
total(9,1)=9
total(10,1)=7
total(11,1)=11
total(12,1)=14

total(1,2)="地理"
total(2,2)="化学"
total(3,2)="历史"
total(4,2)="萝卜课"
total(5,2)="美术"
total(6,2)="体育"
total(7,2)="天文"
total(8,2)="物理"
total(9,2)="音乐"
total(10,2)="英语"
total(11,2)="政治"
total(12,2)="自然"
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style>
TD {	FONT-SIZE: 9pt}
</style></head>
<body topmargin=5 leftmargin=0 scroll=AUTO>
<%call table1(total,200,20,10,20,400,200,"A")%>
<%call table1(total,200,320,10,15,400,250,"B")%>
</body>
</html>


<%
function table1(total,table_x,table_y,thickness,table_width,all_width,all_height,table_type)
'参数含义(传递的数组，横坐标，纵坐标，柱子的厚度，柱子的宽度，图表的宽度，图表的高度,图表的类型)
'纯ASP代码生成图表函数1——柱状图
'作者：龚鸣(Passwordgm) QQ:25968152 MSN:passwordgm@sina.com Email:passwordgm@sina.com
'本人非常愿意和ASP,VML,FLASH的爱好者在HTTP://www.ilisten.cn进行交流和探讨
'非常感谢您使用这个函数，请您使用和转载时保留版权信息，这是对作者工作的最好的尊重。
dim tb_color(14,2)
tb_color(1,1)="#d1ffd1"
tb_color(2,1)="#ffbbbb"
tb_color(3,1)="#ffe3bb"
tb_color(4,1)="#cff4f3"
tb_color(5,1)="#d9d9e5"
tb_color(6,1)="#ffc7ab"
tb_color(7,1)="#ecffb7"
tb_color(8,1)="#d1ffd1"
tb_color(9,1)="#ffbbbb"
tb_color(10,1)="#ffe3bb"
tb_color(11,1)="#cff4f3"
tb_color(12,1)="#d9d9e5"
tb_color(13,1)="#ffc7ab"
tb_color(14,1)="#ecffb7"

tb_color(1,2)="#00ff00"
tb_color(2,2)="#ff0000"
tb_color(3,2)="#ff9900"
tb_color(4,2)="#33cccc"
tb_color(5,2)="#666699"
tb_color(6,2)="#993300"
tb_color(7,2)="#99cc00"
tb_color(8,2)="#00ff00"
tb_color(9,2)="#ff0000"
tb_color(10,2)="#ff9900"
tb_color(11,2)="#33cccc"
tb_color(12,2)="#666699"
tb_color(13,2)="#993300"
tb_color(14,2)="#99cc00"

line_color="#69f"
left_width=70
length=thickness/2
total_no=ubound(total,1)

temp1=0
temp6=0
for i=1 to total_no
	if temp1<total(i,1) then temp1=total(i,1)
	if temp6>total(i,1) then temp6=total(i,1)
next
'if temp6=1 then
'else'出现值小于0的情况
temp1=int(temp1)
if temp6>=0 then
	temp6=int(temp6)
	if temp6>10 then
		temp2=mid(cstr(temp6),2,1)
		if temp2>4 then 
			temp3=(int(temp6/(10^(len(cstr(temp6))-1)))-1)*10^(len(cstr(temp6))-1)
		else
			temp3=(int(temp6/(10^(len(cstr(temp6))-1)))-0.5)*10^(len(cstr(temp6))-1)
		end if
		temp6=temp3
	else
		temp6=0
	end if
'	if temp6-10<0 then temp6=0 else temp6=temp6-10
else
	temp6=int(0-temp6)
	if temp6>10 then
		temp2=mid(cstr(temp6),2,1)
		if temp2>4 then 
			temp3=(int(temp6/(10^(len(cstr(temp6))-1)))+1)*10^(len(cstr(temp6))-1)
		else
			temp3=(int(temp6/(10^(len(cstr(temp6))-1)))+0.5)*10^(len(cstr(temp6))-1)
		end if
		temp6=0-temp3
	else
		temp6=-10
	end if
end if
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
temp4=temp3
response.write "<v:rect id='_x0000_s1027' alt='' style='position:absolute;left:"&table_x+left_width&"px;top:"&table_y&"px;width:"&all_width&"px;height:"&all_height&"px;z-index:-1' fillcolor='#9cf' stroked='f'><v:fill rotate='t' angle='-45' focus='100%' type='gradient'/></v:rect>"
select case table_type
case "A"
	for i=0 to all_height step all_height/5
		if i<>all_height then response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length-i&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height-length-i&"px' strokecolor='"&line_color&"'/>"
		if i<>all_height then response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height-length-i&"px' to='"&table_x+left_width+length&"px,"&table_y+all_height-i&"px' strokecolor='"&line_color&"'/>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+(left_width-15)&"px,"&table_y+i&"px' to='"&table_x+left_width&"px,"&table_y+i&"px'/>"
		response.write ""
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+i&"px;width:"&left_width&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
		temp4=temp4-(temp3-temp6)/5
	next
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y&"px' to='"&table_x+left_width+length&"px,"&table_y+all_height-length&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width-50&"px;top:"&table_y-20&"px;width:100px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>纵坐标</td></tr></table></v:textbox></v:shape>"
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width&"px;top:"&table_y+all_height-9&"px;width:100px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left'>横坐标</td></tr></table></v:textbox></v:shape>"

	line0=(temp3/(temp3-temp6))*all_height
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:2' from='"&table_x+left_width&"px,"&table_y+line0&"px' to='"&table_x+all_width+left_width&"px,"&table_y+line0&"px'></v:line>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+line0-length&"px' to='"&table_x+left_width+length&"px,"&table_y+line0&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+line0-length&"px' to='"&table_x+all_width+left_width&"px,"&table_y+line0-length&"px' strokecolor='"&line_color&"'/>"


'	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+(left_width-15)&"px,"&table_y+line0&"px' to='"&table_x+left_width&"px,"&table_y+line0&"px'/>"
'	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x-15&"px;top:"&table_y+line0-10&"px;width:"&left_width&"px;height:18px;z-index:1'>"
'	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>0</td></tr></table></v:textbox></v:shape>"

	table_space=(all_width-table_width*total_no)/total_no
for i=1 to total_no
	temp_space=table_x+left_width+table_space/2+table_space*(i-1)+table_width*(i-1)
	response.write "<v:rect id='_x0000_s1025' alt='' style='position:absolute;"
	response.write "left:"&temp_space&"px;top:"
	if total(i,1)>0 then
		response.write table_y+line0-(total(i,1)/(temp3-temp6))*all_height
	else
		response.write table_y+line0
	end if
	response.write "px;width:"&table_width&"px;height:"&all_height*(abs(total(i,1))/(temp3-temp6))
	response.write "px;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
	response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' type='gradient'/>"
	response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
	response.write "</v:rect>"

	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&temp_space&"px;top:"
	if total(i,1)>0 then
		response.write table_y+line0-(total(i,1)/(temp3-temp6))*all_height-table_width
	else
		response.write table_y+line0-table_width
	end if
	response.write "px;width:"&table_space+15&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,1)&"</td></tr></table></v:textbox></v:shape>"

	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&temp_space-table_space/2&"px;top:"&table_y+all_height+1&"px;width:"&table_space+table_width&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,2)&"</td></tr></table></v:textbox></v:shape>"
next

Case "B"
	table_space=(all_height-table_width*total_no)/total_no
	temp4=temp6

	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+left_width&"px,"&table_y+all_height+15&"px'/>"
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width-left_width&"px;top:"&table_y+all_height&"px;width:"&left_width&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
	temp4=temp6+(temp3-temp6)/5
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length&"px' to='"&table_x+left_width+all_width&"px,"&table_y+all_height-length&"px' strokecolor='"&line_color&"'/>"
for i=0 to all_width-1 step all_width/5
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+i&"px,"&table_y+all_height-length&"px' to='"&table_x+left_width+length+i&"px,"&table_y+all_height&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length+i&"px,"&table_y+all_height-length&"px' to='"&table_x+left_width+length+i&"px,"&table_y&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+i+all_width/5&"px,"&table_y+all_height&"px' to='"&table_x+left_width+i+all_width/5&"px,"&table_y+all_height+15&"px'/>"
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+i+all_width/5-left_width&"px;top:"&table_y+all_height&"px;width:"&left_width&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
	temp4=temp4+(temp3-temp6)/5
next

for i=1 to total_no
	temp_space=table_space/2+table_space*(i-1)+table_width*(i-1)
	if total(i,1)>=0 then
		response.write "<v:rect id='_x0000_s1025' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)&"px;top:"&table_y+temp_space&"px;width:"&all_width*temp3/(temp3-temp6)*(total(i,1)/temp3)&"px;height:"&table_width&"px;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
		response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' angle='-90' focus='100%' type='gradient'/>"
		response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
		response.write "</v:rect>"
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)+all_width*temp3/(temp3-temp6)*(total(i,1)/temp3)+thickness/2&"px;top:"&table_y+temp_space&"px;width:"&table_space+15&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,1)&"</td></tr></table></v:textbox></v:shape>"
	else
		response.write "<v:rect id='_x0000_s1025' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)-all_width*temp3/(temp3-temp6)*abs(total(i,1)/temp3)&"px;top:"&table_y+temp_space&"px;width:"&all_width*temp3/(temp3-temp6)*abs(total(i,1)/temp3)&"px;height:"&table_width&"px;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
		response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' angle='-90' focus='100%' type='gradient'/>"
		response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
		response.write "</v:rect>"
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)+thickness/2&"px;top:"&table_y+temp_space-table_space/4&"px;width:"&table_space+30&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,1)&"</td></tr></table></v:textbox></v:shape>"
	end if

	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+temp_space&"px;width:"&left_width&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&total(i,2)&"</td></tr></table></v:textbox></v:shape>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)&"px,"&table_y&"px' to='"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)&"px,"&table_y+all_height&"px'></v:line>"
next

case else
end select
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height&"px'><v:stroke endarrow='block'/></v:line>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y&"px' to='"&table_x+left_width&"px,"&table_y+all_height&"px'><v:stroke endarrow='block'/></v:line>"
'end if

end function
%>




