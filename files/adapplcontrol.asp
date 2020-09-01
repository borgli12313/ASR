<% option explicit

	dim mobj, rs, pageAction, i , msg, retVal , delVal, delVal1
	Dim deptList, objList
%>
<html>
<head>
<Title>ASR - Assign IT Account Manager</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>

<!-- #include file="links.asp" -->
<%

	pageAction = Request("pageAction")
	delVal  = Request("delVal")
	delVal1  = Request("delVal1")
	If struser="" Then %>
		<script>
		alert("You have to be in the BAXSIN domain to access this page. Please contact your IT/Administrator.");
		window.close();
		</script>
	<%
	End If ' -- If strUser="" Then -- '

	Set objList = Server.CreateObject("ASRTrans.clsList")
	deptList = objList.PopulateDept()
	'dt= objList.RetrieveDate
	Set objList = nothing
%>


<script language="Javascript" src="fnlist.js"></script>

<script>
var popWin;
function showPopUp(skey) {
	//if (popwin != null) { popWin.close() }
	var url = "poplist.asp" + "?skey=" + skey;
	popWin = open(url, "popupwin", "toolbar=no,scrollbars=yes,width=480,height=200,top=250, left=200");
}
function returnCustomer(customer) {
	frmAdmin.customer.value = customer;
}

function returnAppl(appl) {
	frmAdmin.applname.value = appl;
}

function newDets()
{
	var url = "adapplcontroledit.asp?pageAction=NEW";
	popWin = open(url, "NewApplication", "toolbar=yes,resizable=yes,scrollbars=no,width=600,height=300,top=60, left=150");
}
function GetDets() {
	frmAdmin.pageAction.value = "OK";
	frmAdmin.submit();
}
function showDets(cname, applname)
{
	var url = "adapplcontroledit.asp?pageAction=EDIT&cname=" + escape(cname) + "&applname=" + applname;
	popWin = open(url, "EditApplication", "toolbar=yes,resizable=yes,scrollbars=no,width=600,height=300,top=60, left=100");
}
function DeleteRow(i) {
	if (confirm("Do you want to delete the customer+application?"))
		{
			frmAdmin.pageAction.value = "DEL";
			frmAdmin.delVal.value = frmAdmin.cname(i).value;
			frmAdmin.delVal1.value = frmAdmin.appl(i).value;
			//alert(frmAdmin.delVal.value);
			frmAdmin.submit();
		}
}
function SubmitForm() {

	frmAdmin.pageAction.value = "SEARCH";
	frmAdmin.submit();
}
</script>
<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post" name=frmAdmin  action="adapplcontrol.asp">
<input type="hidden" name="pageAction" value="<%=pageAction%>">
<input type="hidden" name="delVal" value="<%=delVal%>">
<input type="hidden" name="delVal1" value="<%=delVal1%>">
<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Search IT Account Manager</td>
			<td class="rpthr" align="right">
				<input name="cmd2" type="button" id="cmd2" value="New"  onClick="javascript:newDets();">&nbsp;&nbsp;
			  	<input name="cmd2" type="button" id="cmd2" value="Search"  onClick="javascript:SubmitForm();">&nbsp;&nbsp;
			</td>
	    </tr>
	    </table>
	    </td></tr>
	   <tr><td>

		<table width="500" border="0" cellspacing="2" cellpadding="5">


		  <tr>
			  <td width="148" class="lbl1">Customer</td>
			  <td>
			    <input type="text" name="customer" value="<%=Request("customer")%>"> <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('CUSTOMER');">
			  </td>
			</tr>
			<tr>
			  <td width="148" class="lbl1">Application</td>
			  <td>
			    <input type="text" name="applname" value="<%=Request("applname")%>"> <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPL');">
			  </td>
			</tr>
		</table>
		</td></tr>
</table>
<BR>
<table>
	<tr>
		<td>
		<%
		If pageAction  = "SEARCH" then
			Set mobj = Server.CreateObject("ASRMaster.clsApplControl")
			Set rs = mobj.RetrieveSearch(Request("customer"), Request("applname"))

			If rs.eof=false Then%>
					<table border="1" cellpadding="2" cellspacing="0" >
						<tr>
							<td class="trHdr">DEL</td>
							<td class="trHdr">Customer</td>
							<td class="trHdr">Application</td>
							<td class="trHdr">AppMgr</td>
							<td class="trHdr">AppBkpMgr</td>
							<td class="trHdr">Status</td>
						</tr>
						<INPUT type="hidden" name="cname" value="">
						<INPUT type="hidden" name="appl" value="">
						<%
						i=0
						'CName, ApplName, AppMgr, AppBkpMgr, Status
						Do while rs.eof=false %>

						<tr>
						<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=i+1%>);" >
						<img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>

						<INPUT type="hidden" name="cname" value="<%=rs.fields("cName").value%>">
						<INPUT type="hidden" name="appl" value="<%=rs.fields("ApplName").value%>">
						<td> &nbsp;<a href="javascript:showDets('<%=rs.fields(0)%>','<%=rs.fields(1)%>');"><%=rs.fields("cname").value%> </a></td>
						<td> &nbsp;<%=rs.fields("ApplName").value%> </td>
						<td> &nbsp;<%=rs.fields("AppMgr").value%> </td>
						<td> &nbsp;<%=rs.fields("AppBkpMgr").value%> </td>
						<td> &nbsp;<%=rs.fields("Status").value%> </td>
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
		Set mobj = Server.CreateObject("ASRMaster.clsApplControl")
		retVal =  mobj.DeleteRecord(Request("delVal"), Request("delVal1"))

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
				ElseIf retVal = "ACExists" then
				%>
					<script>
					alert("You cannot delete this application. The Account manager is assigned to this application.");
					</script>
				<%

				ElseIf retVal = "ReqExists" then
				%>
					<script>
					alert("You cannot delete this application. Requests have been created using this customer.");
					</script>
				<%Else
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