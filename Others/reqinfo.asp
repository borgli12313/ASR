<%@ Language=VBScript %>
<% dim pageAction, reqno 
pageAction = Request("pageAction")
reqno = Request("reqno")
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

Your request has been created. Request number is <%=Request("reqno")%>
Click the OK button to proceed.
<form method="post" name=reqinfo>
<INPUT id=submit1 type=submit value=OK name="OK">
<input type="hidden" name="pageAction" value="SEARCH">
<INPUT type="hidden" value="<%=Request("reqno")%>" name=reqno></P>
</form>
<%
	If (pageAction = "SEARCH") then
		Response.Redirect "search.asp"
	End If
%>

</BODY>
</HTML>
