/*
	Menu Javascript v 0.9
	Author: ECdeisgn.net
	Create Date: 2004.3.10
	Modify Date: 2004.3.10

*/

var gsBodyId = "";
var gsWorkFrame = "";
var gsCurSubmenuId = "";
var gsIcon_Path = "";
var gsIcon_Open = "";
var gsIcon_Close = "";
var goDir_Open = "";
var goDir_Close = "";



function TurnSubmenu () {
	oDiv = document.getElementById (this.id+".box");
	oImg = document.getElementById (this.id+".image");
	if (oDiv.style.display == "block") {
		oDiv.style.display = "none";
		oImg.src = gsIcon_Close;
	} else {
		oDiv.style.display = "block";
		oImg.src = gsIcon_Open;
	}
	ResetWidth ();
}

function MenuMouseover () {
	this.className = "menu_down";
}

function MenuMouseout () {
	this.className = "menu";
}

function SubmenuMouseover () {
	if (this.id != gsCurSubmenuId)
		this.className = "submenu_down";
}

function SubmenuMouseout () {
	if (this.id != gsCurSubmenuId)
	this.className = "submenu";
}

function RunPrg () {
	if (gsCurSubmenuId) {
		oDiv = document.getElementById (gsCurSubmenuId);
		oDiv.className = "submenu";
	}
		
	gsCurSubmenuId = this.id;
	this.className = "submenu_active";
	oFrame = parent.document.getElementById (gsWorkFrame);
	oFrame.src = this.PrgPath;
}

function ResetWidth () {
	var oDivs = document.getElementsByTagName("DIV");
	for (var i = 0; i < oDivs.length; i++) {
		oMenu = document.getElementById (oDivs[i].id);
		oMenu.style.width = document.body.clientWidth;
	}
}

function NoActive () {
	return false;
}


/* CreateMenu 
	aData is 2 dim MENU array:
	[x][0] 層 1->第一層 2->第二層
	   [1] 選單名稱
	   [2] 程式路徑
*/   
function CreateMenu (aData) {
	var oBody, oDiv, oDiv2;
	var nMenuCnt = 0;
	var nSubmenuCnt = 0;
	var i = 0;
	var j = 0;
	
	// Preload images
	goDir_Open = new Image();
	goDir_Close = new Image();
	goDir_Open.src = gsIcon_Open;
	goDir_Close.src = gsIcon_Close;
	
	oBody = document.getElementById (gsBodyId);
	for (i = 0 ; i < aData.length ; i++) {
		oDiv = document.createElement ("DIV");			// Create Menu DIV
		oSpan = document.createElement ("SPAN");
		if (aData[i][0] == "1") { 						// Menu
			if (oDiv2) {								// Menu_box
				oBody.appendChild (oDiv2);
			}
			nMenuCnt += 1;
			nSubmenuCnt = 0;
			oDiv.id = "menu"+nMenuCnt;
			oDiv.className = "menu";
			
			oImg = document.createElement ("IMG");
			oImg.id = "menu"+nMenuCnt+".image";
			oImg.style.width = "16px";
			oImg.style.height = "16px";
			oImg.align = "top";
			oImg.src = gsIcon_Close;
			oDiv.appendChild (oImg);
			oDiv.innerHTML=oDiv.innerHTML+'&nbsp;';

			oSpan.innerHTML = aData[i][1];
			oSpan.style.width = '120px';
			oDiv.appendChild (oSpan);
			oDiv.title=aData[i][1];
			
			oDiv.onclick = TurnSubmenu;
			oDiv.onmouseover = MenuMouseover;
			oDiv.onmouseout = MenuMouseout;
			oBody.appendChild (oDiv);
			
			oDiv2 = document.createElement ("DIV");		// Create Menu_box
			oDiv2.id = "menu"+nMenuCnt+".box";
			oDiv2.style.display = "none";				// Hidden DIV
			oDiv2.style.overflow = "hidden";
		} else {										// Submenu
			nSubmenuCnt += 1;
			oDiv.id = "submenu"+nMenuCnt+"."+nSubmenuCnt;
			oDiv.className = "submenu";

			oImg = document.createElement ("IMG");
			oImg.id = "menu"+nMenuCnt+".image";
			oImg.style.width = "16px";
			oImg.style.height = "16px";
			oImg.align = "top";
			oImg.src = gsIcon_Path+aData[i][3];
			oDiv.appendChild (oImg);
			oDiv.innerHTML=oDiv.innerHTML+'&nbsp;';

			oSpan.innerHTML = aData[i][1];
			oSpan.style.width = '120px';
			oDiv.appendChild (oSpan);
			oDiv.title=aData[i][1];

			oDiv.PrgPath = aData[i][2];					// 自定物件變數 PrgPath 紀錄程式路徑
			oDiv.onclick = RunPrg;
			oDiv.onmouseover = SubmenuMouseover;
			oDiv.onmouseout = SubmenuMouseout;
			oDiv2.appendChild (oDiv);
		}
	}
	if (oDiv2) {
		oBody.appendChild (oDiv2);
	}
}	
//document.body.style.MozUserSelect='none';
