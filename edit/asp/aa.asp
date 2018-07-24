<%
upload_url1=split(Request.ServerVariables("Path_Info"),"/edit/asp/aa.asp")
upload_url=upload_url1(0)&"/Uploadfile/pic/"

response.write upload_url
response.end
%>