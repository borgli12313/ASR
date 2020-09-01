<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Reports</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->

<% dim mobj, rs, pageAction 
	pageAction = Request.Form("pageAction")
	
	%>
<script>	
function GetDets() {
if (frmRep.reqno.value == "")
		{
		alert("Please enter the request number ");
        frmRep.reqno.focus();
        return false;
		}	

	frmRep.pageAction.value = "OK";
	frmRep.submit();
}

</script>
<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post"  id=form1 name=frmRep action="rptreqnostatus.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">

	<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr>
			<td class="rpthr" colspan=2 align="middle">Status Change Report</td>
	    </tr>
	    <tr>
			<td width="100" class="lbl1">Request No</td>
			<td><input name="reqno" style="WIDTH: 100px; HEIGHT: 20px" size=37 value="<%=Request("ReqNo")%>" >
			&nbsp;<INPUT id=cmd1 type=button value=OK name=cmd1 style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:GetDets();"></td>
	    </tr>

	</table>
<br>

	<%
	If pageAction  = "OK" then
	
	
		Set mobj = Server.CreateObject("ASRTrans.clsReport")
		Set rs=mobj.RetrieveReqNoStatusReport(Request("ReqNo"))		
		%>

  		<table width="500" border="1" cellpadding="1" cellspacing="1" >
			<tr>
				<th class="lbl1"> Status</td>
				<th class="lbl1"> User</td>			
				<th class="lbl1"> Date</td>			
			</tr>
		<%	Do While rs.eof=false %>
				<tr>
					<td class="">&nbsp;<%=rs.fields(0).value%>&nbsp;</td>
					<td class="" >&nbsp;<%=rs.fields(1).value%>&nbsp;</td>
					<td class="" >&nbsp;<%=rs.fields(2).value%>&nbsp;</td>
				</tr>
		<%
				rs.movenext
			loop
	
		rs.close
	
		Set rs = nothing
		Set mobj = nothing
	End If
	%>
	</table>

</body>
</html>
