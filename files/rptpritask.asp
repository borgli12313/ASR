<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->

<% 

dim mobj, rs, i 

%>


<form method="post"  id=form1 name=frmRep action="rptbacklog.asp">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr>
			<td class="rpthr" align="middle">Prioritized Tasks</td>
	    </tr>
	</table>
<br>

	<%	
		Set mobj = Server.CreateObject("ASRTrans.clsReport")
		Set rs=mobj.RetrieveBacklogReport()		
		
		%>

  		<table width="900" border="1" cellpadding="1" cellspacing="1" >
			<tr>
			<th class="lbl1" width="50"> <%=rs.fields(0).name%></td>
			<th class="lbl1" width="50"> <%=rs.fields(12).name%>
				<BR><BR> <%=rs.fields(8).name%></td>
			<th class="lbl1" width="100"> <%=rs.fields(1).name%></td>
			<th class="lbl1" width="270"> <%=rs.fields(2).name%>
				<BR><br><%=rs.fields(5).name%></td>
			<th class="lbl1" width="130"> <%=rs.fields(3).name%></td>
			<th class="lbl1" width="50"> <%=rs.fields(4).name%></td>
			
			<th class="lbl1" width="50"> <%=rs.fields(6).name%>
				<br><%=rs.fields(7).name%></td>
			<th class="lbl1" width="150"> <%=rs.fields(9).name%>
				<br><br><%=rs.fields(10).name%></td>
			<th class="lbl1" width="150"> <%=rs.fields(11).name%></td>			

			</tr>
			
		<%	Do While rs.eof=false %>
				<tr>
				<td width="50" valign="top">&nbsp; <%=rs.fields(0).value%>&nbsp;</td>
				<td width="50" valign="top">&nbsp;<%=rs.fields(12).value%>&nbsp;
					<br><br>&nbsp;<%=rs.fields(8).value%>&nbsp;</td>
				<td width="100" valign="top">&nbsp;<%=rs.fields(1).value%>&nbsp;</td>
				<td width="270" valign="top">&nbsp;<%=rs.fields(2).value%>&nbsp;
					<br><br>&nbsp;<%=trim(rs.fields(5).value)%>&nbsp;</td>
				<td width="130" valign="top">&nbsp;<%=rs.fields(3).value%>&nbsp;</td>
				<td width="50" valign="top">&nbsp;<%=rs.fields(4).value%>&nbsp;</td>
				
				<td width="50" valign="top">&nbsp;<%=rs.fields(6).value%>&nbsp;
					<BR><br>&nbsp;<%=rs.fields(7).value%>&nbsp;</td>
				
				<td width="150" valign="top">&nbsp;<%=rs.fields(9).value%>&nbsp;
					<br><br>&nbsp;<%=rs.fields(10).value%>&nbsp;</td>
				<td width="150" valign="top">&nbsp;<%=rs.fields(11).value%>&nbsp;</td>
					
				</tr>
		<%
				rs.movenext
			loop
	
		rs.close
	
		Set rs = nothing
		Set mobj = nothing
	%>
	</table>

</body>
</html>
