<%
OPtion Explicit
Dim strUser


	strUser=Request.ServerVariables("LOGON_USER")
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)
	response.write struser
%>