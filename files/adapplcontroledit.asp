<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	Dim mobj, strUser, retVal, msg, pageAction
	Dim deptList, objList

	strUser=Request.ServerVariables ("LOGON_USER")
	'strUser = "hema"
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)
	pageAction = Request("pageAction")

	Set mobj = Server.CreateObject("ASRMaster.clsApplControl")

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
						if (opener.document.title=="ASR - Assign IT Account Manager")
						{
						opener.document.forms(0).pageAction.value = "SEARCH";
						opener.document.forms(0).submit();
						}
					}
					catch(e) {
					    alert("The account manager has been assigned.");
					    document.location="admin.asp"
					}
					window.close();
					</script>
				<%ElseIf retVal = "Exists" then
				%>
					<script>
					alert("The account manager has been already assigned.");
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
						if (opener.document.title=="ASR - Assign IT Account Manager")
						{
						opener.document.forms(0).pageAction.value = "SEARCH";
						opener.document.forms(0).submit();
						}
					}
					catch(e) {
					    alert("The account manager details has been modified.");
					    document.location="admin.asp"
					}
					window.close();
					</script>
				<%
					Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "EDIT" then
			mobj.RetrieveRecord Request("cname"), Request("applname")
		End If '--pageAction = ADD --'

	End If '-- Access Failed --'
%>

<html>
<head>
<Title>ASR - Assign IT Account Manager </Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">

</head>

<script language="Javascript" src="fnlist.js"></script>
<script>
var popWin;
function showPopUp(skey) {
	//if (popwin != null) { popWin.close() }
	var url = "poplist.asp" + "?skey=" + skey;
	popWin = open(url, "popupwin", "toolbar=no,scrollbars=yes,width=480,height=300,top=250, left=200");
}

function returnCustomer(customer) {
	frmAdd.cname.value = customer;
}
function returnAppl(applname) {
	frmAdd.applname.value = applname;
}
function returnAppMgr(AppMgr)
{
	frmAdd.appmgr.value = AppMgr;
}
function returnAppBkpMgr(AppMgr)
{
	frmAdd.appbkpmgr.value = AppMgr;
}
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
	if(trim(frmobj.applname.value) == "")
		{
		alert("Please enter the application name ");
        frmobj.applname.focus();
        return false;
		}
	if(trim(frmobj.appmgr.value) == "")
		{
		alert("Please enter the Account Manager name ");
        frmobj.appmgr.focus();
        return false;
		}
	if(trim(frmobj.appbkpmgr.value) == "")
		{
		alert("Please enter the Back up Account Manager name ");
        frmobj.apbkppmgr.focus();
        return false;
		}
		
	if(svalue=='EDIT')    {
	document.frmAdd.pageAction.value = "UPD";	}	if(svalue=='NEW')    {    document.frmAdd.pageAction.value = "ADD";    }
	frmobj.submit();

}
</script>
<body >

<% if msg <> "" then %><div class="msginfo"><%= msg %></div><% end if %>
<br>
<form method="post"  id=form1 name=frmAdd  action="adapplcontroledit.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="UID" value="<%=strUser%> ">

	<br><!-- r1 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Assign IT Account Manager </td>
			<td class="trHdr" align=right>
			<% If pageAction = "NEW" or pageAction = "ADD" Then%>				<input type="button" value="OK" name="save" onClick="javascript:submitForm('NEW');">
			<%Else%>				<input type="button" value="OK" name="save" onClick="javascript:submitForm('EDIT');">
			<% End If%>
			</td>
		</tr>
		</table>
	</td></tr>

	<tr><td>
	<table width="550" border="0" cellspacing="2" cellpadding="5">

		<tr>
		  <td width="148" class="lbl1">Customer</td>
		  <td>
		    <input type="text" name="cname" value="<%=Request("cname")%>"  readonly >
		    <%If pageAction = "NEW" or pageAction = "ADD"  then%>
		    <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('CUSTOMER');" >
		    <%End If%>
		  </td>
		</tr>
		<tr>
		  <td width="148" class="lbl1">Application</td>
		  <td>
		    <input type="text" name="applname" value="<%=Request("applname")%>"  readonly  >
		    <%If pageAction = "NEW" or pageAction = "ADD"  then%>
		    <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPL');">
		    <%End If%>
		  </td>
		</tr>

		<tr>
      <td width="148" class="lbl1">IT Account Manager</td>
      <td>
        <input type="text" name="appmgr" value="<%= mobj.AppMgr%>">
        <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPMGR');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Backp up IT Account Manager</td>
      <td>
        <input type="text" name="appbkpmgr" value="<%=mobj.AppBkpMgr %>">
        <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPBKPMGR');">
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