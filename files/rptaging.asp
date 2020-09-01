<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->

<% dim mobj, rs, i, introw  %>
<br>
	<table width="500" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
		
		<tr>
			<td class="rpthr" align="middle" colspan=2><b>Aging Report</b></td>
	    </tr>
		 	<tr height=20>
	<td width=80> </td>	 
	<td > </td>
</tr>
	<%
	Set mobj = Server.CreateObject("ASRTrans.clsReport")
	Set rs=mobj.RetrieveAgingReportGrp()
	%>
	<tr>
	<td width=80> </td>
	<td>
  	<table  border="1" cellpadding="1" cellspacing="1" >
		<tr>  
			<th class="trHdr">&nbsp;NoOfDays&nbsp;</th>
			<th class="trHdr">&nbsp;NoOfRequests&nbsp;</th> 
		</tr>
	<%	Do While rs.eof=false %>
			<tr> 
			<td class="" width=150 align="center">&nbsp;<%= rs.fields(1).value %> Days &nbsp;</td>
			<td class="" width=50 align="center">&nbsp;<%= rs.fields(2).value %>&nbsp;</td>				 
			</tr>
	<% 
			rs.movenext
		loop %> 
	<%rs.close
	
	'dets
	Set rs=mobj.RetrieveAgingReport()	
	introw=1 
	%>
	</table> 
	</td></tr>
	</table>
<br>
<br>
<br>
  	<table width="800" border="1" cellpadding="1" cellspacing="1" >
		<tr> 			
		<th class="trHdr">&nbsp;SNO&nbsp;</th>
			<% for i=0 to rs.fields.Count-1 %>
				<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
			<% next %>
		</tr>
	<%	Do While rs.eof=false %>
			<tr>
			<td class="">&nbsp;<%= introw %>&nbsp;</td>
			 <% for i=0 to rs.fields.Count-1 %>
				<td class="">&nbsp;<%= rs.fields(i).value %>&nbsp;</td>
			<% next %>
				 
			</tr>
	<% 
			introw = introw + 1
			rs.movenext
		loop %>  
	<%rs.close
	
	Set rs = nothing
	Set mobj = nothing

	%>
	</table>


	


</body>
</html>
