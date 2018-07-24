<!--#include file="conn.asp" -->
<!--#include file="const1.asp" -->
<!--#include file="Inc/Function.asp" -->
<%call tendergotorightstatue

'response.write session("user_id")
'response.end%>
<html>
<head>
<LINK REL=STYLESHEET TYPE="text/css" HREF="css/menu.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
p {font-family: "Arial, Helvetica, sans-serif"; font-size: 11pt}
td {font-family: "Arial, Helvetica, sans-serif"; font-size: 11pt ; line-height:normal; }
A{font-family: "Arial, Helvetica, sans-serif";text-transform: none; text-decoration: none; font-size: 11pt}
select{font-family: "Arial, Helvetica, sans-serif";font-size: 11pt;}
input{font-family: "Arial, Helvetica, sans-serif";font-size: 11pt;}
a:hover {font-family: "Arial, Helvetica, sans-serif";text-decoration:underline; color: #FFFFFF; font-size: 11pt}
body {font-family: "Arial, Helvetica, sans-serif";font-size: 11pt}
div {font-family: "Arial, Helvetica, sans-serif"; font-size: 11pt}
-->
</style>
<title>选单</title>
<base target="main">
</head>
<body ID="Body" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onSelectStart='return false;' bgcolor="#CCCCCC">
<script language="JavaScript1.2" src="js/menu.js"></script>
<script>
document.oncontextmenu=new Function("event.returnValue=false;"); //禁止右键功能,单击右键将无任何反应 

var aData = new Array (0);
gsBodyId = "Body";
gsWorkFrame = "main";
gsIcon_Path = "images/";
gsIcon_Open = "images/dir_open.png";
gsIcon_Close = "images/dir_close.png";
<%
dim menuI
menuI = -1

if session("iflogin")=0 and session("statue")=1 then
%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"公告","menu_gonggao","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"查看公告","gonggao.asp","filenew.png");

aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"竞价信息","menu_addmubiao","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"推荐竞价<%=getcompanycompetitiveinfo(session("user_id"),1)%>","tender_recommend.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"正在竞价<%=getcompanycompetitiveinfo(session("user_id"),2)%>","tender_doing.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"尚未中标<%=getcompanycompetitiveinfo(session("user_id"),3)%>","tender_did.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"中标项目<%=getcompanycompetitiveinfo(session("user_id"),4)%>","tender_get.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"即将竞价<%=getcompanycompetitiveinfo(session("user_id"),5)%>","tender_future.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"我的收藏<%=getcompanycompetitiveinfo(session("user_id"),6)%>","tender_mycollection.asp","filenew.png");

aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"个人信息","menu_userinfo","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"查看资料","company_main_less.asp?id=<%=session("user_id")%>","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"修改密码","changepassword.asp","filenew.png");
<%
elseif session("iflogin")=0 and session("statue")=0 then
%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"公告","menu_gonggao","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"查看公告","gonggao.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"个人信息","menu_userinfo","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"查看资料","company_main_less.asp?id=<%=session("user_id")%>","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"修改资料","company_edit_first.asp","filenew.png");
<%
elseif session("iflogin")>0 then
	if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"竞价管理","menu_checkmubiao","filenew.png");
		<%if session("iflogin")-1=0 or session("iflogin")-99=0 then%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"发布竞价","admin_tender_add.asp","filenew.png");
		<%end if%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"所有竞价","admin_tender_check.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"尚未开始<%=getallcompetitiveinfo(1)%>","admin_tender_future.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"正在竞价<%=getallcompetitiveinfo(2)%>","admin_tender_doing.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"等待定标<%=getallcompetitiveinfo(3)%>","admin_tender_decide.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"完成竞价<%=getallcompetitiveinfo(4)%>","admin_tender_did.asp","filenew.png");

aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"流标管理","menu_checkaoping","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"流标项目<%=getallcompetitiveinfo(5)%>","admin_tender_float.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"删除竞价<%=getallcompetitiveinfo(6)%>","admin_tender_float_del.asp","filenew.png");

aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"统计分析","menu_checkaoping","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"数量折线图","admin_zhexian_year.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"金额折线图","admin_zhexian_money.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"数量饼状图","admin_bing_year.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"金额饼状图","admin_bing_money.asp","filenew.png");
	<%end if%>

aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"个人信息","menu_userinfo","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"密码更改","changepassword.asp","filenew.png");
	<%if session("iflogin")-99=0 then%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"参数配置","menu_userinfo","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"企业类型","admin_companytype.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"资质级别","admin_companygrade.asp","filenew.png");
	<%end if
	if session("iflogin")-1=0 or session("iflogin")-2=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (1,"系统管理","menu_checkmubiao","filenew.png");
		<%if session("iflogin")-99=0 then%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"系统设置","admin_config.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"注册格式","admin_fileprepare.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"注册须知","admin_shouldknow.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"公告管理","admin_gonggao.asp","filenew.png");
		<%end if%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"供应商管理","admin_company.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"供应商审核","admin_company_shen.asp","filenew.png");
		<%if session("iflogin")-99=0 then%>
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"管 理 员","admin_login.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"意见建议","admin_yijian.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"日志管理","admin_rizi.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"数据中心","data_s.asp","filenew.png");
aData [<%menuI=menuI+1:response.write(menuI)%>] = new Array (2,"初 始 化","new.asp","filenew.png");
		<%end if
	end if
end if
%>
CreateMenu (aData);
document.body.onload = ResetWidth;
document.body.onresize = ResetWidth;
window.onresize = ResetWidth;
</script>
</body>
</html>