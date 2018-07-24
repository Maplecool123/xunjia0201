<%
Dim Class_Rs,OptionStyle
Set Class_Rs = conn.execute("Select NoDown_ClassId,NoDown_ClassName,NoDown_Depth,NoDown_Child,NoDown_ParentStr,NoDown_ChildStr from [NoDown_Class] where NoDown_TypeId = " & NoDown_TypeId & " order by NoDown_RootId,NoDown_Orders")
response.write("				<Select name=""ClassId"">" & vbcrlf & "					<option value=""0"">-----∑÷¿‡-----</option>" & vbcrlf)
do while not Class_Rs.eof
	If Class_Rs(2) = 0 Then
		If OptionStyle = "background-color:#F3F3F3;" Then OptionStyle = "background-color:#ffffff;" Else OptionStyle = "background-color:#F3F3F3;"
	End If
		response.write("					<option value=""" & Class_Rs(0) & """" & iif(editid = Class_Rs(0)," Selected=""Selected""","") & " style=""" & OptionStyle & """>")
		If Class_Rs(2) > 0 Then
			for Ci_Class = 1 to Class_Rs(2)
				Response.Write("&nbsp;&nbsp;&nbsp;")
			Next
		End If
		If Class_Rs(3) > 0 Then
			Response.Write("©Ô ")
		Else
			Response.Write("©¿ ")
		End if
	response.write(Class_Rs(1) & "</option>" & vbcrlf)
	Class_Rs.MoveNext 
Loop
Set Class_Rs = Nothing
response.write("				</Select>" & vbcrlf)
%>