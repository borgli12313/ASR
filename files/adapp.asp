<% option explicit

	dim mobj, rs, pageAction, i , msg, retVal 

%>
<html>
<head>
<Title>ASR - Values</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>

<!-- #include file="links.asp" -->
<%
	pageAction = Request("pageAction") 
	If struser="" Then %>
		<script>
		alert("You have to be in the BAXSIN domain to access this page. Please contact your IT/Administrator.");
		window.close();
		</script>
	<%
	End If ' -- If strUser="" Then -- '
%>
<script language="Javascript" src="fnlist.js"></script>
<script> 

function showDets(AppCode)
{
	var url = "adappedit.asp?pageAction=EDIT&AppCode=" + escape(AppCode);
	popWin = open(url, "EditValue", "toolbar=yes,resizable=yes,scrollbars=no,width=500,height=180,top=60, left=100");
}
 
</script>
<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post" name=frmAdmin action="adapp.asp">
<input type="hidden" name="pageAction" value="<%=pageAction%>">
<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Application Values List</td>			
	    </tr>
	    </table>
	    </tr></td>

</table>
<BR>
<table>
	<tr>
		<td>
		<%
		Set mobj = Server.CreateObject("ASRMaster.clsApp") 
			Set rs = mobj.RetrieveList()

			If rs.eof=false Then%>
					<table border="1" cellpadding="2" cellspacing="0" >
						<tr> 
							<td class="trHdr">Name</td>
							<td class="trHdr">Value</td>
						</tr>
						<%
						i=0

						Do while rs.eof=false %>

						<tr>
						<td> &nbsp;<a href="javascript:showDets('<%=rs.fields(0).value%>');"><%=rs.fields(2).value%> </a></td>
						<td> &nbsp;<%=rs.fields(1).value%> </td>
						</tr>
						<%
						rs.movenext
						i=i+1
						loop%>
					<%
			Else %>
				<div class="msginfo">No Details found!</div>
			<%End if 
			rs.close 
			Set rs = nothing  
		Set mobj = nothing  
		%>
			</form>
		</td>
	</tr>
</table>
<BR>
<!-- #include file="footer.asp" -->

</body>
</html>