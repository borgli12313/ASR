
<html>
<head>
<title>ASR - Reports</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->
<form>
	<table width="400" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="400" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Reports</td>
			</tr>
		</table>
	</td></tr>	<tr><td>
		<table width="400" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td>
				<table border="0" cellspacing="2" cellpadding="5">
					<tr>
					 	<td><a href="rptstatus.asp">Status Report</a></td>
					</tr>
					<tr>
					 	<td><a href="rptreqnostatus.asp">Status Change Report (Request Number)</a></td>
					</tr>					<tr>
					 	<td><a href="rptvolstat.asp">Monthly Volume Status Report</a></td>
					</tr>
					<tr>
					 	<td><a href="rptcustvolstat.asp">Customer Volume Status Report</a></td>
					</tr>
					<tr>
					 	<td><a href="rptbaxvolstat.asp">Volume Status Report</a></td>
					</tr>
					<tr>
					 	<td><a href="rptcustdets.asp">Customer Detail Report</a></td>
					</tr>					<tr>
					 	<td><a href="rptperkpi.asp">Performance KPI Report</a></td>
					</tr>  					<tr>
					 	<td><a href="rptqualitykpi.asp">Quality KPI Report</a></td>
					</tr>
					<tr>
					 	<td><a href="rptpritask.asp">Prioritized Tasks</a></td>
					</tr>					<tr>
					 	<td><a href="viewApplControl.asp">IT Account Managers List</a></td>
					</tr>
										
				</table>

			</td>
		</tr>
		</table>	</td></tr>	</table>
	<%If appAccessLevel="1" Then %>
	<table width="400" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="400" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Restricted Reports</td>
			</tr>
		</table>
	</td></tr>	<tr><td>		<table border="0" cellspacing="2" cellpadding="5">			<tr>
			 	<td><a href="rptintcost.asp">Internal Costing Report</a>			 				 	<br>&nbsp;&nbsp;&nbsp;(Charge Sheet Submitted to Finance)
			 	</td>			 	
			</tr>
			<tr>
			 	<td><a href="rptintcostother.asp">Internal Costing Report - Others</a>		 				 	<br>&nbsp;&nbsp;&nbsp;(Non Chargable Division)</td>
			</tr>			<tr>
			 	<td><a href="rptintcostbax.asp">Internal Costing Report - BAX</a>		 				 	<br>&nbsp;&nbsp;&nbsp;(Program name Starting with BAX...)</td>
			</tr>			<tr>
			 	<td><a href="rptaging.asp">Aging Report</a></td>
			</tr>
		</table>			</td></tr>	</table>		  <%End If %>	
</form>
</body>
</html>