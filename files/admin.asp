<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<html>
<head>
<title>ASR - Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>

<!-- #include file="links.asp" -->
<form>
<INPUT type="hidden" name="pageAction" value="<%= pageAction %>">
<INPUT type="hidden" name="UID" value="<%=struser %>">

<table width="400" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="400" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">ASR Admin</td>
			</tr>
		</table>
	</td></tr>
	<tr><td>
	<%If appAccessLevel="1" Then %>
		<table border="0" cellspacing="2" cellpadding="5">
			<tr>
				<td><a href="adcustomer.asp">Customer Maintenance</a></td>
			</tr>
			<tr>
				<td><a href="adappln.asp">Application Maintenance</a></td>
			</tr>
			<tr>
				<td><a href="addivision.asp">Division Maintenance</a></td>
			</tr>			<tr>
				<td><a href="adapp.asp">Project Cost Entry period</a></td>
			</tr> 			<tr>
				<td><a href="adapplcontrol.asp">Assign IT Account Manager</a></td>
			</tr>			<tr>
				<td><a href="adsupporthc.asp">Monthly Headcount-Manhour Entry</a></td>	
			</tr>			
		</table>	  <%End If %>			
	</td></tr>
	</table>
</form>
<!-- #include file="footer.asp" -->
</body>
</html>