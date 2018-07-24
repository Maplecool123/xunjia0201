<%
'分页
Function ShowPage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim CountPage,Ci_ShowPage
	If totalnumber mod maxperpage=0 Then
    	CountPage = totalnumber \ maxperpage
  	Else
    	CountPage = totalnumber \ maxperpage+1
  	End If
	ShowPage = ""
  	ShowPage = ShowPage & vbcrlf & "<div style=""text-align:center;position: inherit;"">" & vbcrlf
	If ShowTotal = true Then 
		ShowPage = ShowPage & "	共&nbsp;<b>" & totalnumber & "</b>&nbsp;" & strUnit & "&nbsp;&nbsp;" & vbcrlf
	End If
	sfilename = JoinChar(sfilename)
  	If CurrentPage < 2 Then
    	ShowPage = ShowPage & "	首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
  	Else
    	ShowPage = ShowPage & "	<a href=""" & sfilename & "page=1"">首页</a>&nbsp;&nbsp;"
    	ShowPage = ShowPage & "<a href=""" & sfilename & "page=" & (CurrentPage-1) & """>上一页</a>&nbsp;&nbsp;"
  	End If

  	If CountPage-currentpage < 1 Then
    		ShowPage = ShowPage & "下一页&nbsp;&nbsp;尾页&nbsp;&nbsp;" & vbcrlf
  	Else
    		ShowPage = ShowPage & "<a href=""" & sfilename & "page=" & (CurrentPage+1) & """>下一页</a>&nbsp;&nbsp;"
    		ShowPage = ShowPage & "<a href=""" & sfilename & "page=" & CountPage & """>尾页</a>&nbsp;&nbsp;" & vbcrlf
  	End If
   	ShowPage = ShowPage & "	页次：<strong><font color=red>" & CurrentPage & "</font>/" & CountPage & "</strong>页&nbsp;&nbsp;" & vbcrlf
    ShowPage = ShowPage & "	<b>" & maxperpage & "</b>" & strUnit & "/页&nbsp;&nbsp;" & vbcrlf
	If ShowAllPages = True Then
		ShowPage = ShowPage & "	转到：<select name=""page"" size=""1"" onchange=""javascript:window.location='" & sfilename & "page=" & "'+this.options[this.selectedIndex].value;"">" & vbcrlf 
    	For Ci_ShowPage = 1 to CountPage   
    		ShowPage=ShowPage & "		<option value=""" & Ci_ShowPage & """"
				If cint(CurrentPage) = cint(Ci_ShowPage) Then ShowPage = ShowPage & " selected=""selected"""
			ShowPage = ShowPage & ">第" & Ci_ShowPage & "页</option>" & vbcrlf  
	    Next
		ShowPage=ShowPage & "	</select>" & vbcrlf
	End If
	ShowPage = ShowPage & "</div>" & vbcrlf
	response.write ShowPage
End Function

'分页时判断在地址栏加入?还是&
Function JoinChar(strUrl)
	If strUrl="" Then
		JoinChar=""
		exit function
	End If
	If InStr(strUrl,"?")<len(strUrl) Then 
		If InStr(strUrl,"?")>1 Then
			If InStr(strUrl,"&")<len(strUrl) Then 
				JoinChar=strUrl & "&"
			Else
				JoinChar=strUrl
			End If
		Else
			JoinChar=strUrl & "?"
		End If
	Else
		JoinChar=strUrl
	End If
End Function
%>