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
	Dim totNo, totActMH

	
	pageAction = Request.Form("pageAction")
%>
<script language="javascript" src="datepicker.js"></script>
<script language="Javascript" src="fnlist.js"></script>
<script>	

var popWin;
function showPopUp(skey) {
	var url = "poplist.asp" + "?skey=" + skey ;
	popWin = open(url, "popupwin", "toolbar=no,scrollbars=yes,width=480,height=300,top=250, left=200");
}

function returnCustomer(customer) 
{
	frmRep.customer.value = customer;
}

function GetDets() {

	if((frmRep.customer.value == ""))
		{
		alert("please select the customer name ");
        frmRep.cmdCustomer.focus();
        return false;
		}	
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
<form method="post"  id=form1 name=frmRep action="rptcustdets.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="dtval" value="<%= date() %>">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Customer Details Report</td>
			<td class="rpthr" align="right"><INPUT id=cmd1 type=button value=OK name=cmd1 style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:GetDets();"></td>
	    </tr>
	    </table>
	    </tr></td>
	    <tr><td>
	    <table width="550" border="0" cellspacing="1" cellpadding="5" >
	    <tr>
		  <td width="100" class="lbl1">Customer</td>
		  <td>
		    <input name="customer" id="customer" type="text" value="<%= Request("customer") %>" readOnly> 
		    <input name="cmdCustomer" type="button" value="..." onClick="javascript:showPopUp('CUSTOMER');">
		  </td>
		</tr>
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
			Set rs=mobj.RetrieveCustDets(Request("customer"),Request("DtFrom"),Request("DtTo"))		
			
			If rs.eof=false Then%>
				<table width="900" border="1" cellpadding="1" cellspacing="1" >
				<tr>
					<td class="trHdr"  width="50" align="middle">&nbsp;<%=rs.fields(0).name%>&nbsp;</td>
					<td class="trHdr"  width="50" align="middle">&nbsp;<%=rs.fields(1).name%>&nbsp;
						<br><%=rs.fields(11).name%>&nbsp;</td>
					<td class="trHdr"  width="150" align="middle">&nbsp;<%=rs.fields(2).name%>&nbsp;</td>
					<td class="trHdr"  width="150" align="middle">&nbsp;<%=rs.fields(3).name%>&nbsp;</td>
					<td class="trHdr"  width="270" align="middle">&nbsp;<%=rs.fields(4).name%>&nbsp;
						<br><br><%=rs.fields(6).name%>&nbsp;</td>
					<td class="trHdr"  width="50" align="middle">&nbsp;<%=rs.fields(5).name%>&nbsp;</td>					
					<td class="trHdr"  width="50" align="middle">&nbsp;ActStartDt&nbsp;
						<br>ActComplDt&nbsp;</td>
					<td class="trHdr"  width="50" align="middle">&nbsp;ActCutinDt&nbsp;
						<br><%=rs.fields(10).name%>&nbsp;</td>					
					<td class="trHdr"  width="50" align="middle">&nbsp;<%=rs.fields(12).name%>&nbsp;</td>
				</tr>
			<%	
				totNo = 0
				totActMH=0
				Do While rs.eof=false %>
					<tr>
					<td width="50" align="middle" valign="top">&nbsp;<%=rs.fields(0).value%>&nbsp;</td>
					<td width="50" align="middle" valign="top">&nbsp;<%=rs.fields(1).value%>&nbsp;
						<br><br> &nbsp;<%=rs.fields(11).value%>&nbsp;</td>
					<td width="150" align="middle" valign="top">&nbsp;<%=rs.fields(2).value%>&nbsp;</td>
					<td width="150" align="middle" valign="top">&nbsp;<%=rs.fields(3).value%>&nbsp;</td>
					<td width="270" align="middle" valign="top">&nbsp;<%=rs.fields(4).value%>&nbsp;
						<br><br> &nbsp;<%=rs.fields(6).value%>&nbsp;</td>
					<td width="50" align="middle" valign="top">&nbsp;<%=rs.fields(5).value%>&nbsp;</td>
					
					<td width="50" align="middle" valign="top">&nbsp;<%=rs.fields(7).value%>&nbsp;
						<br>&nbsp;<%=rs.fields(8).value%>&nbsp;</td>
					<td width="50" align="middle" valign="top">&nbsp;<%=rs.fields(9).value%>&nbsp;
						<br>&nbsp;<%=rs.fields(10).value%>&nbsp;</td>					
					<td width="50" align="middle" valign="top">&nbsp;<%=rs.fields(12).value%>&nbsp;</td>
					
					
					</tr>
			<%
					totNo = totNo + 1
					totActMH = totActMH +  rs.fields(10).value
					rs.movenext
					
				loop%>
				<tr ><td colspan=11>&nbsp;</td></tr >
				
				<tr >
				<td colspan=2><b>Total No. of Requests</b></td> <td><%=totNo%> </td>
				<td colspan=2></td>
				<td colspan=2><b>Total ManHour</b></td> <td><%=totActMH%> </td>
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
	
</form>
</body>
</html>
