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
	Dim PassP()
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
<form method="post"  id=form1 name=frmRep action="rptqualitykpi.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="dtval" value="<%= date() %>">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Quality KPI</td>
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
			Set rs=mobj.RetrieveQualilyKPIReport(Request("DtFrom"),Request("DtTo"))

			count = rs.RecordCount
			ReDim Monthly(count)
			ReDim PassP(count)
			%>
			<table width="" border="0" cellpadding="1" cellspacing="1">
			<tr>
			<td>
			<%
			If rs.eof=false Then
				showgraph = true
			%>
				<table width="" border="1" cellpadding="1" cellspacing="1" >
				<tr>
					<% for i=1 to rs.fields.Count-1 %>
						<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
					<% next %>
				</tr>

			<%
				Monthly(0) = ""
				PassP(0) = 0
				count = 2
				Do While rs.eof=false %>
					<tr>
					<% for i=1 to rs.fields.Count-1 %>
						<td class="" align="right">&nbsp;<%=rs.fields(i).value%>&nbsp;</td>
					<% next %>
					</tr>
			<%		Monthly(count-1) = rs.fields(1).Value
					PassP(count-1) = CDbl(rs.fields(3))/ CDbl(rs.fields(2))*100
					totUAT = totUAT + rs.fields(2).value
					totUATP = totUATP +  rs.fields(3).value
					totUATF = totUATF + rs.fields(4).value
					count = count + 1
					rs.movenext

				loop%>
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

			</td>
			</tr>
			<tr>
			<td>
			<%
			'If showgraph = true Then
				'graph PassP, Monthly
			%>&nbsp;&nbsp;
				<!--<img src="<%=chartURL%>" border="0">-->
			<%
			'End If
			%>
			</td>
			</tr>
			</table>

				<%rs.close
				Set rs=mobj.RetrieveQualilyKPIDets(Request("DtFrom"),Request("DtTo"))

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
				<BR>
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
	%>


</body>
</html>

<%

Sub graph(arrData1, arrLabel)

	Dim cd, data0, data1, data2, data3, labels, c, layer, chartId

	Set cd = CreateObject("ChartDirector.API")

	'The data for the line chart
	data0 = arrData1

	labels = arrLabel

	'Create a XYChart object of size 300 x 180 pixels
	Set c = cd.XYChart(350, 250)

	'Set background color to pale yellow 0xffffc0, with a black edge and a 1
	'pixel 3D border
	Call c.setBackground(&Hffffc0, &H0, 1)

	'Set the plotarea at (45, 35) and of size 240 x 120 pixels, with white
	'background. Turn on both horizontal and vertical grid lines with light
	'grey color (0xc0c0c0)
	Call c.setPlotArea(45, 35, 240, 140, &Hffffff, -1, -1, &Hc0c0c0, -1)

	'Add a legend box at (45, 12) (top of the chart) using horizontal layout
	'and 8 pts Arial font Set the background and border color to Transparent.
	'Call c.addLegend(45, 12, 0, "", 8).setBackground(cd.Transparent)

	'Add a title to the chart using 9 pts Arial Bold/white font. Use a 1 x 2
	'bitmap pattern as the background.
	Call c.addTitle("Quality KPI ", "arialbd.ttf", 9, &Hffffff).setBackground(c.patternColor(Array(&H4000, &H8000), 2))

	'Set the y axis label format to nn%
	Call c.yAxis().setLabelFormat("{value}%")

	'Set the labels on the x axis
	Call c.xAxis().setLabels(labels)

	'Add a line layer to the chart
	Set layer = c.addLineLayer()

	'Add the first line. Plot the points with a 7 pixel square symbol
	Call layer.addDataSet(data0, &Hcf4040, "Pass Percentage").setDataSymbol( cd.SquareSymbol, 7)

	'Enable data label on the data points. Set the label format to nn%.
	Call layer.setDataLabelFormat("{value|2}%")

	'Reserve 10% margin at the top of the plot area during auto-scaling to
	'leave space for the data labels.
	'Call c.yAxis().setAutoScale(1)
	call c.yAxis().setLinearScale(0, 120, 20)

	'Create the image and save it in a session variable
	Session("chart") = c.makeChart2(cd.PNG)
	Session("imgNo") = Session("imgNo") + 1
	chartId = Session.SessionId & "_" & Session("imgNo")
	chartURL = "myimage.asp?img=chart&id=" & chartId


End Sub
%>

