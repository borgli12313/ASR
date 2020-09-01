<%
	Dim strUser 

	strUser=Request.ServerVariables ("LOGON_USER")
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)
%>	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=unicode">
<META content="MSHTML 6.00.2719.2200" name=GENERATOR></HEAD>


<!-- #include file="access2.asp" -->

<BODY>
<table width="500" border="0" cellspacing="5" cellpadding="5">
  <tr align="middle" class="appHdr">&nbsp;&nbsp;Application Service Request (ASR)
  </tr>
  <tr>
        <td colspan=2><font size=2><b>User ID: </b> <%=strUser%></font></td>
         
        <td><font size=1><b><%=Date()%></b></font></td>
  </tr> 
  <tr>
    <td align="middle" class="box1">Home</td>
    <td align="middle" class="box1"><A href="search.asp">Search</A></td>

    <%If appAccessLevel="1" Then %>
    <td align="middle" class="box1"><A href="admin.asp" >Admin</A></td>
    <%End If %>
    <td align="middle" class="box1"><A href="reports.asp" >Reports</A></td>

  </tr>
</table></BODY></HTML>
