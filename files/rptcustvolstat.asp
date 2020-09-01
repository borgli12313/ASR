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
	Dim mobj, rs, pageAction, i
	Dim totOpen, totClose, totReq

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
<form method="post"  id=form1 name=frmRep >
	<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
	<INPUT type="hidden" name="dtval" value="<%= date() %>">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
			<table width="500" border="0" cellspacing="1" cellpadding="5" >
				<tr>
				<td class="rpthr" align="middle">Customer Volume Status Report</td>
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
				<tr>
				<td width="100" class="lbl1">Sort By</td>

				<td><select name="sortby">
					<option value="Customer">Customer </option>
					<option value="TOTAL">TOTAL </option>
					</select>
					<script>
						<% if Request("sortby") <> "" then %>document.forms[0].sortby.value = "<%= Request("sortby") %>"; <% end if %>
					</script>
				</td>
				<td><select name="order">
							<option value="ASC"> Ascending</option>
							<option value="DESC">Descending</option>
				        </select></td>

				<script>
				<% if Request("order") <> "" then %>document.forms[0].order.value = "<%= Request("order") %>"; <% end if %>
				</script>

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
			Set rs=mobj.RetrieveCustVolStatusReport(Request("DtFrom"),Request("DtTo"))

			If rs.eof=false Then
			rs.sort= Request("sortby") & " " & Request("order")

			%>
				<table width="" border="1" cellpadding="1" cellspacing="1" >
				<tr>
					<th class="trHdr">&nbsp;<%= rs.fields(0).name %>&nbsp;</th>
					<th class="trHdr"  width=80>&nbsp;OPEN&nbsp;</th>
					<% for i=2 to rs.fields.Count-1 %>
						<th class="trHdr" width=80>&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
					<% next %>
				</tr>
			<%	Do While rs.eof=false %>
					<tr>
					<% for i=0 to rs.fields.Count-1 %>
						<td class="" align="middle">&nbsp;<%=rs.fields(i).value%>&nbsp;</td>
					<% next %>
					</tr>
			<%
					totOpen = totOpen + rs.fields(1).value
					totClose = totClose +  rs.fields(7).value
					totReq = totReq +  rs.fields(10).value
					rs.movenext

				loop%>
				<tr ><td colspan=11>&nbsp;</td></tr >

				<tr >
				<td colspan=2><b>Total Open Request</b></td> <td align="middle"><%=totOpen%> </td>
				<td></td>
				<td colspan=2><b>Total Closed Request</b></td> <td align="middle"><%=totClose%> </td>
				<td></td>
				<td colspan=2><b>Total Request</b></td> <td align="middle"><%=totReq%> </td>
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
