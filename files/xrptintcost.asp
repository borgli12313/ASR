<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->

<% dim mobj, rs, i, totamt, subtot,  strprv, strprvdiv 
DIM subtotdiv1,subtotdiv2,subtotdiv3,subtotdiv4
%>
<br>
	<table width="900" border="0" cellspacing="0" cellpadding="5" bordercolor="#9966CC">
		
		<tr>
			<td class="rpthr" align="middle"><b>INTERNAL COST TRANSER FORM - <%=monthname(month(date())) %></b></td>
	    
	<td align="middle" class="rpthr" ><input type=button name=cmdlogin value=" Open in Excel "
				onclick="javascript:document.location='rptintcostxl.asp'"></td>
	    </tr>
	</table>
		
<br>

	<%
	Set mobj = Server.CreateObject("ASRTrans.clsReport")
	Set rs=mobj.RetrieveIntCost()	
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
						<td align="center" class="lbl1"> &nbsp; <b><%=formatnumber(subtot,2)%></b></td>
						<td > &nbsp;&nbsp;</td>
						</tr>
						
					<%

					subtot=0 
					End if%>
				<tr>
				<%if strprvdiv <> rs.fields("Division").value  then %>
				<td class="">&nbsp;<%= rs.fields("Division").value %>&nbsp;</td>
				<%else%> 
				<td class="">&nbsp;&nbsp;</td>
				<%end if%>
				<td class="">&nbsp;<%= rs.fields("Program").value %>&nbsp;</td>
				<%else%>
				<tr>
				<%if strprvdiv <> rs.fields("Division").value  then %>
				<td class="">&nbsp;<%= rs.fields("Division").value %>&nbsp;</td>
				<%else%> 
				<td class="">&nbsp;&nbsp;</td>
				<%end if%>
				<td class="">&nbsp;&nbsp;</td>
				<%end if%>
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
			<td align="center"> &nbsp; <b><%=formatnumber(subtot,2)%></b></td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td colspan =10> &nbsp; </td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp;<b>SUB TOTAL</b>&nbsp;</td>
			<td class="" colspan =2>&nbsp; LOGISTICS &nbsp;</td>
			<td class="" align="right" colspan =2>&nbsp;<b><%=formatnumber(subtotdiv1,2)%></b>&nbsp;</td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp; &nbsp;</td>
			<td class="" colspan =2 >&nbsp; AIR &nbsp;</td>
			<td class="" align="right"  colspan =2 >&nbsp;<b><%=formatnumber(subtotdiv2,2)%></b>&nbsp;</td>
			<td> &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp; &nbsp;</td>
			<td class="" colspan =2 >&nbsp; OCEAN &nbsp;</td>
			<td class="" align="right" colspan =2  >&nbsp;<b><%=formatnumber(subtotdiv3,2)%></b>&nbsp;</td>
			<td> &nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class=""  align="right" colspan =5>&nbsp; &nbsp;</td>
			<td class="" colspan =2>&nbsp; SUPPORT &nbsp;</td>
			<td class="" align="right" colspan =2>&nbsp;<b><%=formatnumber(subtotdiv4,2)%></b>&nbsp;</td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		
		<tr>
			<td colspan =10> &nbsp; </td>
		</tr>
		<tr> 
			<td class=""  align="right" colspan =7>&nbsp;<b>TOTAL</b>&nbsp;</td>
			<td class="" align="right" colspan =2 >&nbsp;<b><%=formatnumber(totamt,2)%></b>&nbsp;</td>
			<td > &nbsp;&nbsp;</td>
		</tr>
		
		<tr><td></td></tr></table>
	<%rs.close
	
	Set rs = nothing
	Set mobj = nothing

	%>
	


	


</body>
</html>
