<!--
//询问对话框
function hintandturn(str,url,torf)
{
	if(confirm(str) && torf){
		window.location=url;
		return true;
	}else{
		return false;
		}
}

function doCheckDetail(intId,dw,dh){
	showModelessDialog(intId,0,"dialogWidth:"+ dw +"px;dialogHeight:"+ dh +"px;status:no;scroll:yes;");
}

function doCheckDetailEx(intId,dw,dh){
	showModalDialog(intId,0,"dialogWidth:"+ dw +"px;dialogHeight:"+ dh +"px;status:no;scroll:yes;");
}

//选择所有复选框
function CheckAll11(FormName,CheckboxName,chkAllName){
  for (var i=0;i<eval("document." + FormName + "." + CheckboxName + ".length");i++){
		var e = eval("document." + FormName + "." + CheckboxName + "[" + i + "]");
		if (e.Name != chkAllName){
			e.checked = eval("document." + FormName + "." + chkAllName + ".checked");
		}
    }
}
function UnSelectAll(FormName,chkAllName){
    eval("document." + FormName + "." + chkAllName + ".checked=1")	
}


//鼠标控制Div背景颜色
function divsetbg(obj){
	obj.style.backgroundColor = '#ffffff';
}
function divrebg(obj){
	obj.style.backgroundColor = '';
}
function divrebgcheck(obj){
	var objChildCheck = document.all ? obj.children[0].children[0] : obj.childNodes[1].childNodes[1];
	if(objChildCheck.checked){
		return false;
	}
	obj.style.backgroundColor = '';
}


function onloadimg(obj,widthnum,heightnum){  
	//if(obj.width > obj.height){  
	//	if(obj.width > widthnum){  
	//		obj.width = widthnum;  
	//	}  
	//}else{  
	//	if(obj.height > heightnum){  
	//		obj.height = heightnum;  
	//	}  
	//}
	if(obj.width > widthnum){  
		obj.width = widthnum;  
	} 
}  




function ClickFace(FaceId){
	for(var i=1;i<=12;i++){
		eval("document.GB.faceid" + i + ".style.border='1px solid #f0f0f0';");
	}
	eval("document.GB." + FaceId + ".style.border='1px solid #000000';");
	eval("document.GB.GBFace.value=document.GB." + FaceId + ".name;");
}
function SelectFace(FaceType){
	for(var i=1;i<=12;i++){
		eval("document.GB.faceid" + i + ".src='images/gbface/"+ FaceType + i + ".jpg';");
		eval("document.GB.faceid" + i + ".name='"+ FaceType + i + ".jpg';");
	}
	ClickFace("faceid1");
}
function ChkGB(){
	var FormName=document.GB;
	if(FormName.GBTitle.value=="")
	{
		alert("请填写标题再提交表单。");
		FormName.GBTitle.focus();
		return false;
	}
	if(FormName.GBContent.value=="")
	{
		alert("请填写内容再提交表单。");
		FormName.GBContent.focus();
		return false;
	}
	return true;
}







//鼠标停留在有Title或Alt标签元素上时显示框框
tPopWait=50;		//停留tWait毫秒后显示提示。
tPopShow=5000;		//显示tShow毫秒后关闭提示
showPopStep=20;
popOpacity=99;
sPop=null;
curShow=null;
tFadeOut=null;
tFadeIn=null;
tFadeWaiting=null;
document.write("<div id='ncpopLayer' style='position:absolute;z-index:1000;background-color: #ffffff; border: 1px #555555 solid; font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0);'></div>");
function showPopupText(){
var o=event.srcElement;
	MouseX=event.x;
	MouseY=event.y;
	if(o.alt!=null && o.alt!=""){o.ncpop=o.alt;o.alt=""};
        if(o.title!=null && o.title!=""){o.ncpop=o.title;o.title=""};
	if(o.ncpop!=sPop) {
			sPop=o.ncpop;
			clearTimeout(curShow);
			clearTimeout(tFadeOut);
			clearTimeout(tFadeIn);
			clearTimeout(tFadeWaiting);	
			if(sPop==null || sPop=="") {
				ncpopLayer.innerHTML="";
				ncpopLayer.style.filter="Alpha()";
				ncpopLayer.filters.Alpha.opacity=0;	
				}
			else {
				if(o.dyclass!=null) popStyle=o.dyclass 
					else popStyle="cPopText";
				curShow=setTimeout("showIt()",tPopWait);
			}
			
	}
}
function showIt(){
		ncpopLayer.className=popStyle;
		ncpopLayer.innerHTML=sPop;
		popWidth=ncpopLayer.clientWidth;
		popHeight=ncpopLayer.clientHeight;
		if(MouseX+12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24
			else popLeftAdjust=0;
		if(MouseY+12+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24
			else popTopAdjust=0;
		ncpopLayer.style.left=MouseX+12+document.body.scrollLeft+popLeftAdjust;
		ncpopLayer.style.top=MouseY+12+document.body.scrollTop+popTopAdjust;
		ncpopLayer.style.filter="Alpha(Opacity=0)";
		fadeOut();
}
function fadeOut(){
	if(ncpopLayer.filters.Alpha.opacity<popOpacity) {
		ncpopLayer.filters.Alpha.opacity+=showPopStep;
		tFadeOut=setTimeout("fadeOut()",1);
		}
		else {
			ncpopLayer.filters.Alpha.opacity=popOpacity;
			tFadeWaiting=setTimeout("fadeIn()",tPopShow);
			}
}
function fadeIn(){
	if(ncpopLayer.filters.Alpha.opacity>0) {
		ncpopLayer.filters.Alpha.opacity-=1;
		tFadeIn=setTimeout("fadeIn()",1);
		}
}
document.onmouseover=showPopupText;

//-->