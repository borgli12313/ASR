<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->

<% dim mobj, rs, i, totalreq %>

	<table width="300" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
		<tr>
			<td class="rpthr" align="middle">Status Report</td>
	    </tr>
	</table>
<br>

	<%
	Set mobj = Server.CreateObject("ASRTrans.clsReport")
	Set rs=mobj.RetrieveStatusReport()		
	%>

  	<table width="300" border="1" cellpadding="1" cellspacing="1" >
		<tr>
			<th class="lbl1"> Status</td>
			<th class="lbl1"> No of Requests</td>			
			
		</tr>
	<%	Do While rs.eof=false %>
			<tr>
				<td class="">&nbsp;<%=rs.fields(0).value%>&nbsp;</td>
				<td class="" align="middle">&nbsp;<%=rs.fields(1).value%>&nbsp;</td>
			</tr>
	<%
			totalreq= totalreq + rs.fields(1).value
			rs.movenext
		loop %>
		<tr><td></td></tr>
		<tr>
			<td class="">&nbsp;<b>Total</b>&nbsp;</td>
			<td class="" align="middle">&nbsp;<b><%=totalreq%></b>&nbsp;</td>
		</tr>
	<%rs.close
	
	Set rs = nothing
	Set mobj = nothing

	%>
	</table>


	


</body>
</html>
