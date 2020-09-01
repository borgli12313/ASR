<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<script language="Javascript" src="fnlist.js"></script>
<script>	
function GetDets() {
if((frmRep.ipyear.value == ""))
		{
		frmRep.ipyear.value =<%=year(date)%>;
		}	
	frmRep.pageAction.value = "OK";
	frmRep.submit();
}

</script>
<body>
<body>
<!-- #include file="links.asp" -->

<% dim mobj, rs, i, pageAction, totamt, subtot, strprv 
pageAction = Request.Form("pageAction")
%>
<br>
<form method="post"  id=form1 name=frmRep action="rptintcostbax.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<br>
	<table width="900" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
		
		<tr>
			<td class="rpthr" align="middle"><b>INTERNAL COST TRANSER FORM BAX </b></td>			
			<td align="left" class="rpthr" ><input type=button name=cmdlogin value="Open in Excel"
				onclick="javascript:document.location='rptintcostbaxxl.asp?ipyear=<%=request("ipyear")%>&ipmonth=<%=request("ipmonth")%>'">
			&nbsp;&nbsp;&nbsp;<INPUT id=cmd1 type=button value=OK name=cmd1 style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:GetDets();"></td>
	    </tr>
	    </table>
	    <table width="500" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
	
	    <tr> 
	    
			<td width="100" class="lbl1">Year</td>
			<td width="100"><INPUT type="text" name="ipyear" size=4 value="<%=request("ipyear")%>" maxlength="4" onKeyPress="if(!((window.event.keyCode >= '48')&&(window.event.keyCode <= '57'))){alert('Please enter NUMERIC values only.');return false;}"
			></td>
			 
			<td width="100" class="lbl1">Month</td>
			<td> 
			<select name="ipmonth">
						<option value="01">01</option>
						<option value="02">02</option>
						<option value="03">03</option>
						<option value="04">04</option>
						<option value="05">05</option>
						<option value="06">06</option>
						<option value="07">07</option>
						<option value="08">08</option>
						<option value="09">09</option>
						<option value="10">10</option>
						<option value="11">11</option>
						<option value="12">12</option>
					</select>
					<script>SetItemValue("ipmonth","<%=request("ipmonth")%>");</script>
			</td>	
	    </tr>
	</table>
		
<br>

	<%
	If pageAction  = "" then%>
	<script>
		frmRep.ipyear.value =<%=year(date)%>;
		<%if len(month( date())-1)=1 then%>
		SetItemValue("ipmonth","0" + <%=cstr(month(date())-1)%>);
		<%else%>
		SetItemValue("ipmonth",<%=cstr(month(date())-1)%>);		
		<%end If%>
	</script>

<%ElseIf pageAction  = "OK" then  
	Set mobj = Server.CreateObject("ASRTrans.clsReport")
	Set rs=mobj.RetrieveIntCostBAX(Request("ipyear")& Request("ipmonth"))	
	totamt=0	
	subtot=0 
	%>

  	<table width="900" border="1" cellpadding="1" cellspacing="1" >
		<tr> 			
			<% for i=0 to rs.fields.Count-1 %>
				<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
			<% next %>
		</tr>
	<%	Do While rs.eof=false %>
			 
				<%if strprv <> rs.fields("Program").value  then  
					If strprv<>"" then %>
						<tr>
						<td colspan =8> &nbsp;&nbsp;</td>
						<td align="center" class="lbl1"> &nbsp; <b><%=formatnumber(subtot,2)%></b></td>
						<td colspan =2> &nbsp;&nbsp;</td>
						</tr>
						<tr><td colspan =9> &nbsp; </td></tr>
					<%

					subtot=0 
					End if%>
				<tr><td class="">&nbsp;<%= rs.fields("Program").value %>&nbsp;</td>
				<%else%>
				<tr><td class="">&nbsp;&nbsp;</td>
				<%end if%>
				<td class="">&nbsp;<%= rs.fields("DeptName").value %>&nbsp;</td>
				<td class="">&nbsp;<%= rs.fields("Activity").value %>&nbsp;</td>
				<td class="">&nbsp;<%= rs.fields("ASRCount").value %>&nbsp;</td>
				<td class="">&nbsp;<%= rs.fields("Category").value %>&nbsp;</td>
				<td class="" align="right">&nbsp;<%= rs.fields("#of Hrs").value %>&nbsp;</td>
				<td class="" align="center" width="50">&nbsp;<%= rs.fields("Basis").value %>&nbsp;</td>
				<td class="" align="center">&nbsp;<%= rs.fields("ChargeRate").value %>&nbsp;</td>
				<td class="" align="right">&nbsp;<%= formatnumber(rs.fields("Amount").value,2) %>&nbsp;</td>
				<td class="">&nbsp;<%= rs.fields("Details").value %>&nbsp;</td> 
			</tr>
	<% 
			subtot= subtot + formatnumber(rs.fields("Amount").value,2)
			totamt= totamt + formatnumber(rs.fields("Amount").value,2)
			strprv=rs.fields("Program").value 
			rs.movenext
		loop %>
		<tr>
			<td colspan =8> &nbsp;&nbsp;</td>
			<td align="center"> &nbsp; <b><%=formatnumber(subtot,2)%></b></td>
			<td colspan =2> &nbsp;&nbsp;</td>
		</tr>
		<tr><td></td></tr>
		<tr> 
			<td class=""  align="right" colspan =7>&nbsp;<b>TOTAL</b>&nbsp;</td>
			<td class="" align="right" colspan =2 >&nbsp;<b><%=formatnumber(totamt,2)%></b>&nbsp;</td>
			<td > &nbsp;&nbsp;</td>
		</tr>
	<%rs.close
	
	Set rs = nothing
end if
	Set mobj = nothing

	%>
	</table>


	


</body>
</html>
