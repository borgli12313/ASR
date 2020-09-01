<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% option explicit%>

<%
dim mobj, rs, i, pageAction, msg,  totamt, subtot,  strprv, strprvdiv 
DIM subtotdiv1,subtotdiv2,subtotdiv3,subtotdiv4

      Response.Buffer = TRUE
      Response.ContentType = "application/vnd.ms-excel"
 %>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script language="Javascript" src="fnlist.js"></script>
<script>	


</script>
<body> 

<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post"  id=form1 name=frmRep action="rptintcost.asp">
<br>
	<table width="900" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
		
		<tr>
			<td class="rpthr" align="middle" colspan=3><b>INTERNAL COST TRANSER FORM </b></td>
		 </tr>
	    </table>
	    <table width="500" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
	
	    <tr> 
			<td width="100" class="lbl1">Year:  <%=request("ipyear")%> </td>
			 
			<td width="100" class="lbl1">Month:  <%=request("ipmonth")%></td>	
	    </tr>
	</table>
		
<br>

	<%
	
	Set mobj = Server.CreateObject("ASRTrans.clsReport")
	Set rs=mobj.RetrieveIntCostOthers(Request("ipyear")& Request("ipmonth"))	
	totamt=0	
	subtot=0 
	subtotdiv1=0 
	subtotdiv2=0 
	subtotdiv3=0 
	subtotdiv4=0 
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
						<td align="center" class="lbl1"><b><%=formatnumber(subtot,2)%></b></td>
						<td > &nbsp;</td>
						</tr>
						
					<%

					subtot=0 
					End if%>
				<tr>
				<%if strprvdiv <> rs.fields("Division").value  then %>
				<td class=""><%= rs.fields("Division").value %></td>
				<%else%> 
				<td class="">&nbsp;&nbsp;</td>
				<%end if%>
				<td class=""><%= rs.fields("Program").value %></td>
				<%else%>
				<tr>
				<%if strprvdiv <> rs.fields("Division").value  then %>
				<td class=""><%= rs.fields("Division").value %></td>
				<%else%> 
				<td class="">&nbsp;</td>
				<%end if%>
				<td class="">&nbsp;</td>
				<%end if%>
				<td class=""><%= rs.fields("Activity").value %></td>
				<td class=""><%= rs.fields("ASRCount").value %></td>
				<td class=""><%= rs.fields("Category").value %></td>
				<td class="" align="right"><%= rs.fields("#of Hrs").value %></td>
				<td class="" align="center" width="50"><%= rs.fields("Basis").value %></td>
				<td class="" align="center"><%= rs.fields("ChargeRate").value %></td>
				<td class="" align="right"><%=formatnumber(rs.fields("Amount").value,2)%></td>				
				<td class=""><%= rs.fields("Details").value %></td> 
			</tr>
	<% 
			subtot= subtot + formatnumber(rs.fields("Amount").value,2)
			totamt= totamt + formatnumber(rs.fields("Amount").value,2)
			strprv=rs.fields("Program").value 
			strprvdiv =rs.fields("Division").value 
			select case Ucase(rs.fields("Division").value)
			case "LOGISTICS"
				subtotdiv1 = subtotdiv1 + formatnumber(rs.fields("Amount").value,2)
			case "AIR"
				subtotdiv2 = subtotdiv2 + formatnumber(rs.fields("Amount").value,2)
			CASE "OCEAN"
				subtotdiv3 = subtotdiv3 + formatnumber(rs.fields("Amount").value,2)
			CASE "SUPPORT"
				subtotdiv4 = subtotdiv4 + formatnumber(rs.fields("Amount").value,2)
			end Select
			rs.movenext
		loop %>
		<tr>
			<td colspan =8> &nbsp;&nbsp;</td>
			<td align="center"><b><%=formatnumber(subtot,2)%></b></td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td colspan =10> &nbsp; </td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp;<b>SUB TOTAL</b>&nbsp;</td>
			<td class="" colspan =2>&nbsp; LOGISTICS &nbsp;</td>
			<td class="" align="right" colspan =2><b><%=formatnumber(subtotdiv1,2)%></b></td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp; &nbsp;</td>
			<td class="" colspan =2 >&nbsp; AIR &nbsp;</td>
			<td class="" align="right"  colspan =2 ><b><%=formatnumber(subtotdiv2,2)%></b></td>
			<td> &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp; &nbsp;</td>
			<td class="" colspan =2 >&nbsp; OCEAN &nbsp;</td>
			<td class="" align="right" colspan =2  ><b><%=formatnumber(subtotdiv3,2)%></b></td>
			<td> &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp; &nbsp;</td>
			<td class="" colspan =2>&nbsp; SUPPORT &nbsp;</td>
			<td class="" align="right" colspan =2><b><%=formatnumber(subtotdiv4,2)%></b></td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		
		<tr>
			<td colspan =10> &nbsp; </td>
		</tr>
		<tr> 
			<td class=""  align="right" colspan =7>&nbsp;<b>TOTAL</b>&nbsp;</td>
			<td class="" align="right" colspan =2 ><b><%=formatnumber(totamt,2)%></b></td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		
		<tr><td></td></tr></table>
	<%rs.close
	
	Set rs = nothing
	Set mobj = nothing

	%>
	


	
</form>

</body>
</html>
