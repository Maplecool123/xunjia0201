<!--#include file="Conn.asp" -->
<!--#include file="Const.asp" -->
<!--#include file="Inc/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href='css/workspace.css' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="js/Function.js"></script>
<SCRIPT language="JavaScript" src="inc/day.js"></SCRIPT>
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
  
function collapse(img, objName)
{
	var obj;
	obj = document.getElementById(objName);
	if (img.src.indexOf('open') != -1)
	{
		img.src = img.src.replace('open', 'close');
		obj.style.display = 'none';
	}
	else
	{
		img.src = img.src.replace('close', 'open');
		obj.style.display = '';
	}
}

function collapseall(img)
{
	var obj;
	if (img.src.indexOf('open') != -1)
	{
		img.src = img.src.replace('open', 'close');
		for (x=1;x<=<%=maxrecord%>;x++)
		{
			obj = document.getElementById("tender"+x);
			if (obj)
			{
				obj.style.display = 'none';
			}
		}
	}
	else
	{
		img.src = img.src.replace('close', 'open');
		for (x=1;x<=<%=maxrecord%>;x++)
		{
			obj = document.getElementById("tender"+x);
			if (obj)
			{
				obj.style.display = '';
			}
		}		
	}
}

document.oncontextmenu=new Function("event.returnValue=false;"); //禁止右键功能,单击右键将无任何反应 
</SCRIPT> 
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">