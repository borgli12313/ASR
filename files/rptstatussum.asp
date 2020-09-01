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
	dim mobj, rs, pageAction ,i
	Dim totOpen, totClose

	
	pageAction = Request.Form("pageAction")
%>
<script language="javascript" src="datepicker.js"></script>
<script language="Javascript" src="fnlist.js"></script>
<script>	
function GetDets() {
	if (frmRep.dtfrom.value == "")
		{
		frmRep.dtfrom.value ='<%=DateAdd("m", -5, date())%>';
		}
	if (frmRep.dtto.value == "")	
		{
		frmRep.dtto.value ='<%=Date()%>';
		}

	if (isDate(frmRep.dtfrom.value,'dd/MM/yyyy')==false)
		{
		alert("Enter valid Date format (DD/MM/YYYY) ");
		frmRep.dtfrom.focus();
		return false;
		}
	if (isDate(frmRep.dtto.value,'dd/MM/yyyy')==false)
		{
		alert("Enter valid Date format (DD/MM/YYYY) ");
		frmRep.dtto.focus();
		return false;
		}	
	if (compareDates(frmRep.dtto.value,'dd/MM/yyyy',frmRep.dtval.value,'dd/MM/yyyy')==1)
		{
		alert("You cannot enter any future date.");
		frmRep.dtto.focus();
		return false;
		}
	
	
	frmRep.pageAction.value = "OK";
	frmRep.submit();
}

</script>
<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post"  id=form1 name=frmRep action="rptstatussum.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="dtval" value="<%= date() %>">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Status Change Report</td>
			<td class="rpthr" align="right"><INPUT id=cmd1 type=button value=OK name=cmd1 style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:GetDets();"></td>
	    </tr>
	    </table>
	    </tr></td>
	    <tr><td>
	    <table width="500" border="0" cellspacing="1" cellpadding="5" >
	    <tr>
			<td width="100" class="lbl1">From Date</td>
			<td><input name="dtfrom" style="WIDTH: 100px; HEIGHT: 20px" size=37 value="<%=Request("DtFrom")%>" >
				<a href="javascript:show_calendar('frmRep.dtfrom');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
					<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
			</td>
	    
			<td width="100" class="lbl1">To Date</td>
			<td><input name="dtto" style="WIDTH: 100px; HEIGHT: 20px" size=37 value="<%=Request("DtTo")%>" >
				<a href="javascript:show_calendar('frmRep.dtto');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
					<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
			</td>
	    </tr>
		</table>
	    </tr></td>
	</table>
<br>

	<%
	If pageAction  = "OK" then
	
		If DateDiff("m", Request("DtFrom"),Request("DtTo")) > 12 Then
			Response.Write "The date range cannot be more than 12 months. Please change the From Date or To date"
			
		else		
	
			Set mobj = Server.CreateObject("ASRTrans.clsReport")
			Set rs=mobj.RetrieveStatusSumReport(Request("DtFrom"),Request("DtTo"))		
			
			If rs.eof=false Then%>
				<table width="" border="1" cellpadding="1" cellspacing="1" >
				<tr>
					<% for i=1 to rs.fields.Count-1 %>
						<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
					<% next %>
				</tr>
			<%	Do While rs.eof=false %>
					<tr>
					<% for i=1 to rs.fields.Count-1 %>
						<td class="" align="middle">&nbsp;<%=rs.fields(i).value%>&nbsp;</td>
					<% next %>
					</tr>
			<%
					totOpen = totOpen + rs.fields(2).value
					totClose = totClose +  rs.fields(8).value
					rs.movenext
					
				loop%>
				<tr ><td colspan=11>&nbsp;</td></tr >
				
				<tr >
				<td colspan=4><b>Total Open Request</b></td> <td><%=totOpen%> </td>
				<td></td>
				<td colspan=4><b>Total Closed Request</b></td> <td><%=totClose%> </td>
				</tr>
				</table>
			<%Else %>
			No records found!
			<%End if
			rs.close
	
			Set rs = nothing
		End If
		Set mobj = nothing
	End If
	%>
	

</body>
</html>
