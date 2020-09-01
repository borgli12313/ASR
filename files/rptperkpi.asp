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
	Dim totUAT, totUATP, totUATF
	Dim Monthly()
	Dim TotalUAT()
	Dim MetTargetDate()
	Dim count, showgraph,chartURL

	showgraph = false
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
<form method="post"  id=form1 name=frmRep action="rptperkpi.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="dtval" value="<%= date() %>">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Performance KPI Report</td>
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
			Set rs=mobj.RetrievePerKPIReport(Request("DtFrom"),Request("DtTo"))

			count = rs.RecordCount
			ReDim Monthly(count-1)
			ReDim TotalUAT(count-1)
			ReDim MetTargetDate(count-1)

			If rs.eof=false Then%>
				<table width="" border="1" cellpadding="1" cellspacing="1" >
				<tr>
					<% for i=1 to rs.fields.Count-1 %>
						<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
					<% next %>
				</tr>
			<%	count = 1
				Do While rs.eof=false
					showgraph = true
			%>
					<tr>
					<% for i=1 to rs.fields.Count-1 %>
						<td class="" align="right">&nbsp;<%=rs.fields(i).value%>&nbsp;</td>

					<% next %>
					</tr>
			<%
					If rs.fields(1).Value<>"" then
						Monthly(count-1) = rs.fields(1).Value
						TotalUAT(count-1) = rs.fields(2).Value - rs.fields(3).Value
						MetTargetDate(count-1) = rs.fields(3).Value
					end if
					totUAT = totUAT + rs.fields(2).value
					totUATP = totUATP +  rs.fields(3).value
					totUATF = totUATF + rs.fields(4).value
					count = count + 1
					rs.movenext

				loop
				%>
				<tr ><td colspan=11>&nbsp;</td></tr >

				<tr >
				<td ><b>Total </b></td>
				<td align="right"><%=totUAT%> </td>
				<td align="right"><%=totUATP%> </td>
				<td align="right"><%=totUATF%> </td>
				<td align="right">
				<script>
				document.write(RoundOff(<%=totUATP*100/totUAT%>,2));
				</script>
				 </td>
				</tr>
				</table>
				<br>

			<%
			'If showgraph = true Then
				'graph TotalUAT, MetTargetDate, Monthly
			%>
				&nbsp;&nbsp;
				<!--<img src="<%=chartURL%>" border="0">-->
			<%
			'End If
			%>


				<%rs.close
				Set rs=mobj.RetrievePerKPIDets(Request("DtFrom"),Request("DtTo"))

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
							<td>&nbsp;<%=rs.fields(i).value%>&nbsp;</td>
						<% next %>
						</tr>
				<%
						rs.movenext

					loop%>
					</table>
				<%Else %>
				No Details found!
				<%End if
			Else %>
			No records found!
			<%End if
			rs.close

			Set rs = nothing
		End If
		Set mobj = nothing
	End If



Sub graph(TotalUAT, MetTargetDate, Monthly)

	Dim data0, data1, data2
	Dim labels, cd, c
	Dim chartId, imageMap

	'Create the graphic'
	Set cd = CreateObject("ChartDirector.API")

	'The data for the bar chart'

	data0 = MetTargetDate
	data1 = TotalUAT

	'The labels for the bar chart'
	labels = Monthly

	'Create a XYChart object of size 500 x 320 pixels'
	Set c = cd.XYChart(500, 320)

	'Set the plot area at (45, 25) and of size 239 x 180. Use two alternative'
	'background colors (0xffffc0 and 0xffffe0)'
	Call c.setPlotArea(90, 40, 250, 180).setBackground(&Hffffc0, &Hffffe0)

	'Add a legend box
	'transparent background
	Call c.addLegend(380, 40)

Call c.addTitle("Performance KPI ", "arialbd.ttf", 9, &Hffffff).setBackground(c.patternColor(Array(&H4000, &H8000), 2))
	'Set the labels on the x axis'
	Call c.xAxis().setLabels(labels)

	'Add a stacked bar layer and set the layer 3D depth to 8 pixels'
	Set layer = c.addBarLayer2(cd.Stack, 8)

	'Add three bar layers, each representing one data set. The bars are draw in semi-transparent colors.
Call layer.addDataSet(data0, &Hff8080, "MetTargetDate")
Call layer.addDataSet(data1, &H80ff80, "Exceeded")

'Enable bar label for the whole bar
Call layer.setAggregateLabelStyle()

'Enable bar label for each segment of the stacked bar
Call layer.setDataLabelStyle()

	'Create the image and save it in a session variable
	Session("chart") = c.makeChart2(cd.PNG)
	Session("imgNo") = Session("imgNo") + 1
	chartId = Session.SessionId & "_" & Session("imgNo")
	chartURL = "myimage.asp?img=chart&id=" & chartId

End Sub
	%>


</body>
</html>
