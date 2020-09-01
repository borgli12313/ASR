<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% 
	Dim mobj, strUser, retVal, msg, pageAction 

	strUser=Request.ServerVariables ("LOGON_USER")
	'strUser = "hema"
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)

	Set mobj = Server.CreateObject("ASRMaster.clsApplication")

	If strUser="" then
		msg = "Access Denied"
	Else
		pageAction = Request("pageAction") 
		If pageAction = "ADD" then
			mobj.SetValues Request
			On Error Resume Next
			retVal = mobj.InsertRecord
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then %>

					<script>
					
						    //alert("The new application has been created.");
						var parentWin = opener;
						try
						{
							if (opener.document.title=="ASR - Application Maintenance")
							{
							opener.document.forms(0).pageAction.value = "SEARCH";
							opener.document.forms(0).submit();
							}
						}
						catch(e) {
						    alert("The new application has been created.");
						    document.location="search.asp"
						}
						window.close();
					</script>
				<%ElseIf retVal = "Exists" then
				%>
					<script>
					alert("The application already exists.");
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
						if (opener.document.title=="ASR - Application Maintenance")
						{
						opener.document.forms(0).pageAction.value = "SEARCH";
						opener.document.forms(0).submit();
						}
					}
					catch(e) {
					    alert("The appliction details has been modified.");
					    document.location="appln.asp"
					}
					window.close();
					</script>
				<%
					Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "EDIT" then 
			mobj.RetrieveRecord(Request("applname"))
		End If '--pageAction = ADD --'
	End If '-- Access Failed --'

	

%>

<html>
<head>
<Title>ASR - New Application </Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">

</head>

<script language="Javascript" src="fnlist.js"></script>
<script>

//--SUBMIT FORM---
function submitForm(svalue)
{
	var frmobj = document.forms[0];
	if(trim(frmobj.applname.value) == "")
		{
		alert("Please enter the application name ");
	    frmobj.applname.focus();
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
<form method="post"  id=form1 name=frmAdd action="adapplnedit.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="UID" value="<%=strUser%> ">

	<br><!-- r1 start-->
  <table width="450" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="450" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Application </td>
			<td class="trHdr" align="right">
			<% If pageAction = "NEW" or pageAction = "ADD" Then%>				<input type="button" value="OK" name="save" onClick="javascript:submitForm('NEW');">
			<%Else%>				<input type="button" value="OK" name="save" onClick="javascript:submitForm('EDIT');">
			<% End If%>
		</td>
			</tr>
		</table>
	</td></tr>

	<tr><td>
		<table width="450" border="0" cellspacing="2" cellpadding="5">
			<tr>
			<td width="148" class="lbl1">Application Name</td>
		  <td><input name="applname" id="applname" maxlength=50 style="WIDTH: 263px; HEIGHT: 20px" value="<%= mobj.applname %>" size=37 >
			  
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