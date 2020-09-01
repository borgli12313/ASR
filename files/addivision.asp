<% option explicit

	dim mobj, rs, pageAction, i , msg, retVal ,  delVal

%>
<html>
<head>
<Title>ASR - Division Maintenance</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>

<!-- #include file="links.asp" -->
<%

	pageAction = Request("pageAction")
	delVal  = Request("delVal")
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

function newDets()
{
	var url = "addivisionedit.asp?pageAction=NEW";
	popWin = open(url, "Division", "toolbar=yes,resizable=yes,scrollbars=no,width=500,height=180,top=60, left=150");
}
function GetDets() {
	frmAdmin.pageAction.value = "OK";
	frmAdmin.submit();
}

function showDets(divname)
{
	var url = "addivisionedit.asp?pageAction=EDIT&divname=" + escape(divname);
	popWin = open(url, "EditDivision", "toolbar=yes,resizable=yes,scrollbars=no,width=500,height=180,top=60, left=100");
}
function DeleteRow(i) {
	if (confirm("Do you want to delete the division?"))
		{
			frmAdmin.pageAction.value = "DEL";
			frmAdmin.delVal.value = frmAdmin.divname(i).value;
			frmAdmin.submit();
		}

}
function SubmitForm() {

	frmAdmin.pageAction.value = "SEARCH";
	frmAdmin.submit();
}
</script>
<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post" name=frmAdmin action="addivision.asp">
<input type="hidden" name="pageAction" value="<%=pageAction%>">

<input type="hidden" name="delVal" value="<%=delVal%>">
<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Search Division</td>
			<td class="rpthr" align="right">
				<input name="cmd2" type="button" id="cmd2" value="New"  onClick="javascript:newDets();">&nbsp;&nbsp;
			  	<input name="cmd2" type="button" id="cmd2" value="Search"  onClick="javascript:SubmitForm();">&nbsp;&nbsp;
			</td>
	    </tr>
	    </table>
	    </tr></td>

</table>
<BR>
<table>
	<tr>
		<td>
		<%
		Set mobj = Server.CreateObject("ASRMaster.clsDivision")
		If pageAction  = "SEARCH" then
			Set rs = mobj.RetrieveList()

			If rs.eof=false Then%>
					<table border="1" cellpadding="2" cellspacing="0" >
						<tr>
							<td class="trHdr">DEL</td>
							<td class="trHdr">Division</td>
							<td class="trHdr">Type</td>
						</tr>
						<INPUT type="hidden" name="divname" value="">
						<%
						i=0

						Do while rs.eof=false %>

						<tr>
						<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=i+1%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>

						<INPUT type="hidden" name="divname" value="<%=rs.fields("DivName").value%>">
						<td> &nbsp;<a href="javascript:showDets('<%=rs.fields(0)%>');"><%=rs.fields("DivName").value%> </a></td>
						<td> &nbsp;<%=rs.fields("DivType").value%> </td>
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
		End If
	If pageAction  = "DEL" then
		retVal =  mobj.DeleteRecord(Request("delVal"))
			If Err Then
				msg = "There is some error in deleting your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then
				%>
					<script>
					document.frmAdmin.pageAction.value = "SEARCH";
					document.frmAdmin.submit();
					</script>
				<%
				ElseIf retVal = "Exists" then
				%>
					<script>
					alert("You cannot delete this application. The division is used in the customer.");					
					document.frmAdmin.pageAction.value = "SEARCH";
					document.frmAdmin.submit();
					</script>
				<%
				Else
						msg = "There is some error in deleting your details please try again!"
				End If
			End if
	End if
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