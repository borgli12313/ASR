<% option explicit

	dim mobj, rs, pageAction, i , msg, retVal, delVal, delVal1
dim datefr, dateto 
Set mobj = Server.CreateObject("ASRTrans.clsList")
dateto = mobj.RetrieveCostEntryDtTo()
datefr = mobj.RetrieveCostEntryDtFrom()
%>
<html>
<head>
<Title>ASR - Assign Support Head count</Title>
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
function returnAppMgr(AppMgr)
{
	frmAdmin.itstaff.value = AppMgr;
}
function newDets()
{
	var url = "adsupporthcedit.asp?pageAction=NEW";
	popWin = open(url, "NewSupportHC", "toolbar=yes,resizable=yes,scrollbars=no,width=600,height=300,top=60, left=150");
}
function GetDets() {
	frmAdmin.pageAction.value = "OK";
	frmAdmin.submit();
}
function showDets(cname, itsfaff)
{
	var url = "adsupporthcedit.asp?pageAction=EDIT&cname=" + escape(cname) + "&itstaff=" + itsfaff;
	popWin = open(url, "EditSupportHC", "toolbar=yes,resizable=yes,scrollbars=no,width=600,height=300,top=60, left=100");
}
function DeleteRow(i) {
	if (confirm("Do you want to delete this Support Head count?"))
		{
			frmAdmin.pageAction.value = "DEL";
			frmAdmin.delVal.value = frmAdmin.cname(i).value;
			frmAdmin.delVal1.value = frmAdmin.itstaffname(i).value;
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
<form method="post" name=frmAdmin  action="adsupporthc.asp">
<input type="hidden" name="pageAction" value="<%=pageAction%>">
<input type="hidden" name="delVal" value="<%=delVal%>">
<input type="hidden" name="delVal1" value="<%=delVal1%>">
<table width="500" border="1" cellspacing="1" cellpadding="5" bordercolor="#9966CC">
		<tr><td>
		<table width="500" border="0" cellspacing="1" cellpadding="5" >
		<tr>
			<td class="rpthr" align="middle">Search Support Head count</td>
			<td class="rpthr" align="right">			<%If day(date)>=datefr and day(date)<=dateto Then  %>
				<input name="cmd2" type="button" id="cmd2" value="New"  onClick="javascript:newDets();">&nbsp;&nbsp;
			 <%End If %>	
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
			    <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('CUSTOMER');">
			  </td>
			</tr>
			<tr>
			  <td width="148" class="lbl1">IT Staff</td>
			  <td>
			    <input type="text" name="itstaff" value="<%=Request("itstaff")%>">
			    <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPMGR');">
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
			Set mobj = Server.CreateObject("ASRMaster.clsSupportHC")
			Set rs = mobj.RetrieveSearch(Request("customer"), Request("itstaff"))

			If rs.eof=false Then%>
					<table border="1" cellpadding="2" cellspacing="0" >
						<tr>
							<td class="trHdr">DEL</td>
							<td class="trHdr">Customer</td>
							<td class="trHdr">ITStaff</td>
							<td class="trHdr">HCPercent</td>
						</tr>
						<INPUT type="hidden" name="cname" value="">
						<INPUT type="hidden" name="itstaffname" value="">
						<%
						i=0
						'CName, ApplName, AppMgr, AppBkpMgr, Status
						Do while rs.eof=false %>

						<tr>
						<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=i+1%>);" >
						<img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>

						<INPUT type="hidden" name="cname" value="<%=rs.fields("cName").value%>">
						<INPUT type="hidden" name="itstaffname" value="<%=rs.fields("ITStaff").value%>">
						<td> &nbsp;<a href="javascript:showDets('<%=rs.fields(0)%>','<%=rs.fields(1)%>');"><%=rs.fields("cname").value%> </a></td>
						<td> &nbsp;<%=rs.fields("ITStaff").value%> </td>
						<td> &nbsp;<%=rs.fields("HCPercent").value%> </td>
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
		Set mobj = Server.CreateObject("ASRMaster.clsSupportHC")
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