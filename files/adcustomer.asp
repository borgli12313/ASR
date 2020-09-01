<% option explicit

	dim mobj, rs, pageAction, i , msg, retVal, delVal
	Dim deptList, objList
%>
<html>
<head>
<Title>ASR - Customer Maintenance</Title>
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

	Set objList = Server.CreateObject("ASRTrans.clsList")
	deptList = objList.PopulateDept()
	'dt= objList.RetrieveDate
	Set objList = nothing
%>


<script language="Javascript" src="fnlist.js"></script>

<script>

function newDets()
{
	var url = "adcustomeredit.asp?pageAction=NEW";
	popWin = open(url, "NewCustomer", "toolbar=yes,resizable=yes,scrollbars=no,width=600,height=300,top=60, left=150");
}
function GetDets() {
	frmAdmin.pageAction.value = "OK";
	frmAdmin.submit();
}
function showDets(cname)
{
	var url = "adcustomeredit.asp?pageAction=EDIT&cname=" + escape(cname);
	popWin = open(url, "EditCustomer", "toolbar=yes,resizable=yes,scrollbars=no,width=600,height=300,top=60, left=100");
}
function DeleteRow(i) {
	if (confirm("Do you want to delete the customer?"))
		{
			frmAdmin.pageAction.value = "DEL";
			frmAdmin.delVal.value = frmAdmin.cname(i).value;
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
<form method="post" name=frmAdmin  action="adcustomer.asp">
<input type="hidden" name="pageAction" value="<%=pageAction%>">

<input type="hidden" name="delVal" value="<%=delVal%>">
<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Search Customer</td>
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
		      <input type="text" name="customer" value="<%=Request("customer")%>">
		    </td>
		  </tr>
		  <tr>
			<td width="148" class="lbl1">Division</td>
			<td>
				<select name="dept">
					<option value=""></option>
					<option value="Logistics" >Logistics</option>
					<option value="Ocean">Ocean</option>
					<option value="Air">Air</option>
					<option value="Support">Support</option>
				</select>
				<script>SetItemValue("dept","<%=Request("dept")%>");</script>
			</td>
		</tr>
		  <tr>
			<td width="200" class="lbl1">Customer Type:</td>
			<td>
				<select name="ctype">
					<option value=""></option>
					<option value="N">Normal</option>
					<option value="M">Multiclient AC</option>
					<option value="O">Multiclient NonAC</option>
					<option value="WA">Multiclient West A</option>
					<option value="WD">Multiclient West D</option>
					<option value="MAM">Multiclient (AC - MegaHub)</option>
					<option value="MAC">Multiclient (AC - CLC)</option>
				</select>
				<script>SetItemValue("ctype","<%=Request("ctype") %>");</script>
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
			Set mobj = Server.CreateObject("ASRMaster.clsCustomer")
			Set rs = mobj.RetrieveSearch(Request("customer"), Request("dept"), Request("ctype"))

			If rs.eof=false Then%>
					<table border="1" cellpadding="2" cellspacing="0" >
						<tr>
							<td class="trHdr">DEL</td>
							<td class="trHdr">Customer</td>
							<td class="trHdr">Division</td>
							<td class="trHdr">Status</td>
							<td class="trHdr">Type</td>
						</tr>
						<INPUT type="hidden" name="cname" value="">
						<%
						i=0

						Do while rs.eof=false %>

						<tr>
						<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=i+1%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>

						<INPUT type="hidden" name="cname" value="<%=rs.fields("cName").value%>">
						<td> &nbsp;<a href="javascript:showDets('<%=rs.fields(0)%>');"><%=rs.fields("cname").value%> </a></td>
						<td> &nbsp;<%=rs.fields("DeptName").value%> </td>
						<td> &nbsp;<%=rs.fields("Status").value%> </td>
						<td> &nbsp;<%=rs.fields(4).value%> </td>
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
		Set mobj = Server.CreateObject("ASRMaster.clsCustomer")
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
				ElseIf retVal = "ACExists" then
				%>
					<script>
					alert("You cannot delete this application. The Account manager is assigned to this customer.");
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