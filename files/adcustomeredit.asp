<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	Dim mobj, strUser, retVal, msg, pageAction
	Dim deptList, objList

	strUser=Request.ServerVariables ("LOGON_USER")
	'strUser = "hema"
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)
	pageAction = Request("pageAction")

	Set mobj = Server.CreateObject("ASRMaster.clsCustomer")

	If strUser="" then
		msg = "Access Denied"
	Else

		If pageAction = "ADD" then
			mobj.SetValues Request
			On Error Resume Next
			retVal = mobj.InsertRecord
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then %>
					<script>
					var parentWin = opener;
					try
					{
						if (opener.document.title=="ASR - Customer Maintenance")
						{
						opener.document.forms(0).pageAction.value = "SEARCH";
						opener.document.forms(0).submit();
						}
					}
					catch(e) {
					    alert("The new customer has been created.");
					    document.location="admin.asp"
					}
					window.close();
					</script>
				<%ElseIf retVal = "Exists" then
				%>
					<script>
					alert("The customer already exists.");
					</script>
				<%
					Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "UPD" then
			mobj.SetValues Request
			On Error Resume Next
			retVal = mobj.UpdateRecord
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then %>
					<script>
					var parentWin = opener;
					try
					{
						if (opener.document.title=="ASR - Customer Maintenance")
						{
						opener.document.forms(0).pageAction.value = "SEARCH";
						opener.document.forms(0).submit();
						}
					}
					catch(e) {
					    alert("The customer details has been modified.");
					    document.location="search.asp"
					}
					window.close();
					</script>
				<%
					Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "EDIT" then
			mobj.RetrieveRecord(Request("cname"))
		End If '--pageAction = ADD --'
	Set objList = Server.CreateObject("ASRTrans.clsList")
	deptList = objList.PopulateDept()
	'dt= objList.RetrieveDate
	Set objList = nothing
	End If '-- Access Failed --'
%>

<html>
<head>
<Title>ASR - Customer </Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">

</head>

<script language="Javascript" src="fnlist.js"></script>
<script>

//--SUBMIT FORM---
function submitForm(svalue)
{
	var frmobj = document.forms[0]; 
	
	if(trim(frmobj.cname.value) == "")
		{
		alert("Please enter the Customer name ");
	    frmobj.cname.focus();
	    return false;
		}
	if(svalue=='EDIT')
    {
	document.frmAdd.pageAction.value = "UPD";
	}
	if(svalue=='NEW')
    {
    document.frmAdd.pageAction.value = "ADD";
    }
	frmobj.submit();

}
</script>
<body >

<% if msg <> "" then %><div class="msginfo"><%= msg %></div><% end if %>
<br>
<form method="post"  id=form1 name=frmAdd  action="adcustomeredit.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="UID" value="<%=strUser%> ">

	<br><!-- r1 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Customer </td>
			<td class="trHdr" align=right>
			<% If pageAction = "NEW" or pageAction = "ADD" Then%>
				<input type="button" value="OK" name="save" onClick="javascript:submitForm('NEW');">
			<%Else%>
				<input type="button" value="OK" name="save" onClick="javascript:submitForm('EDIT');">
			<% End If%></td>
			
		</tr>
		</table>
	</td></tr>

	<tr><td>
	<table width="550" border="0" cellspacing="2" cellpadding="5">
		<tr>
		<td width="148" class="lbl1">Customer Name</td>
		<td><input name="cname" id="cname" maxlength=50 style="WIDTH: 263px; HEIGHT: 20px" value="<%= mobj.cname %>" size=37
		<%If pageAction = "EDIT" then%> readonly <%End If%> ></td>
		</tr>
		
		 <tr>
		<td width="148" class="lbl1">Division</td>
		<td>
			<select name="dept">
				<option value="Logistics" selected >Logistics</option>
				<option value="Ocean">Ocean</option>
				<option value="Air">Air</option>
				<option value="Support">Support</option>
				<option value="Sales">Sales</option>
			</select>
			<script>SetItemValue("dept","<%=mobj.DeptName%>");</script>
		</td>
		</tr>
		<tr>
		<td width="148" class="lbl1">Remarks</td>
		<td><input name="remarks" id="remarks" maxlength=100 style="WIDTH: 263px; HEIGHT: 20px" value="<%= mobj.Remarks %>" size=37 > </td>
		</tr>
		<tr>
			<td width="200" class="lbl1">Customer Type:</td>
			<td>
				<select name="ctype">
					<option value="N" selected >Normal</option>
					<option value="M">Multiclient AC</option>
					<option value="O">Multiclient NonAC</option>
					<option value="WA">Multiclient West A</option>
					<option value="WD">Multiclient West D</option>
					<option value="MAM">Multiclient (AC - MegaHub)</option>
					<option value="MAC">Multiclient (AC - CLC)</option>
				</select>
				<script>SetItemValue("ctype","<%=mobj.ctype%>");</script>
			</td>
		</tr>
		<tr>
			<td width="200" class="lbl1">Status:</td>
			<td>
				<select name="status">
					<option value="A" selected >Active</option>
					<option value="I">Inactive</option>
				</select>
				<script>SetItemValue("status","<%=mobj.status%>");</script>
			</td>
		</tr>
  </table>

</td></tr>
</table><!-- r1 end-->

</form>
</body>
</html>
<%
Set mobj = nothing
%>