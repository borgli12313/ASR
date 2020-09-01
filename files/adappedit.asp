<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% 
	Dim mobj, strUser, retVal, msg, pageAction 

	strUser=Request.ServerVariables ("LOGON_USER")
	'strUser = "hema"
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)

	Set mobj = Server.CreateObject("ASRMaster.clsApp")

	If strUser="" then
		msg = "Access Denied"
	Else
		pageAction = Request("pageAction") 
		If pageAction = "UPD" then
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
						if (opener.document.title=="ASR - Values")
						{
						opener.document.forms(0).pageAction.value = "SEARCH";
						opener.document.forms(0).submit();
						}
					}
					catch(e) {
					    alert("The details has been modified.");
					    document.location="adapp.asp"
					}
					window.close();
					</script>
				<%
					Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "EDIT" then 
			mobj.RetrieveRecord(Request("AppCode"))
		End If '--pageAction = ADD --'
	End If '-- Access Failed --'

	

%>

<html>
<head>
<Title>ASR - Values </Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">

</head>

<script language="Javascript" src="fnlist.js"></script>
<script>

//--SUBMIT FORM---
function submitForm(svalue)
{
	var frmobj = document.forms[0];
	
	if((trim(frmobj.AppData.value) == "") && (trim(frmobj.AppData.value) == "0"))
		{
		alert("Please enter the value");
	    frmobj.AppData.focus();
	    return false;
		}	
	if(trim(frmobj.AppData.value) > 31)
		{
		alert("Please enter the value (1 - 31)");
	    frmobj.AppData.focus();
	    return false;
		}	
	if(svalue=='EDIT')    {
	document.frmAdd.pageAction.value = "UPD";	} 
	frmobj.submit();
}
</script>
<body >

<% if msg <> "" then %><div class="msginfo"><%= msg %></div><% end if %>
<br>
<form method="post"  id=form1 name=frmAdd action="adappedit.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="UID" value="<%=strUser%> ">

	<br><!-- r1 start-->
  <table width="450" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="450" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">AppValues</td>
			<td class="trHdr" align="right">
			<input type="button" value="OK" name="save" onClick="javascript:submitForm('EDIT');">
		</td>
			</tr>
		</table>
	</td></tr>

	<tr><td>
		<table width="450" border="0" cellspacing="2" cellpadding="5">
			<tr>
			<td width="148" class="lbl1">Code</td>
		  <td><input type=hidden name="AppCode" id="AppCode" maxlength=50 style="WIDTH: 263px; HEIGHT: 20px" value="<%= mobj.AppCode %>" size=37 >
			  <%= mobj.Remarks %>
		  </td>
		</tr>
		 <tr>
			<td width="200" class="lbl1">Value:</td>
			<td><input name="AppData" value="<%=mobj.AppData%>" style="WIDTH: 116px; HEIGHT: 20px" size=16
			maxlength="4" onKeyPress="if(!((window.event.keyCode >= '48')&&(window.event.keyCode <= '57'))){alert('Please enter NUMERIC values only.');return false;}">
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