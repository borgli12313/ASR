<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	Dim clsReqDet, strUser, retVal
	Dim mobj, rs, deptList, msg, dt, introw
	Dim i, m, arrcount, pageAction, delFile, delIndex

	strUser=Request.ServerVariables ("LOGON_USER")
	'strUser = "hema"

	Set clsReqDet = Server.CreateObject("ASRTrans.clsRequestDets")

	If strUser="" then
		msg = "Access Denied"
	Else
		pageAction = Request.Form("pageAction")
		delFile = Request.Form("delFile")
		delIndex = Request.Form("delIndex")
		if delIndex="" then delIndex=-1
		If pageAction <> "" then clsReqDet.SetValues Request
		If pageAction = "ADD" then
			'clsReqDet.SetValues Request
			On Error Resume Next


			retVal = clsReqDet.InsertRecord
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then
					'Response.Redirect "reqinfo.asp?reqno=" &  clsReqDet.ReqNo
				%>

					<script>
					var parentWin = opener;

						    try
						    {
								if (opener.document.title=="ASR - Search")
								{
								opener.document.forms(0).pageAction.value = "SEARCH";
								opener.document.forms(0).reqno.value =<%=clsReqDet.ReqNo%>;
								opener.document.forms(0).submit();
								}
								window.close();

						    }
						    catch(e) {
						        alert("Your Request has been created. The destination (ASR - SEARCH) window has been closed. Please open the ASR Search screen to modify the request details");
						        document.location="search.asp"
						    }


					</script>
				<%
					Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "DEL" then
			'clsReqDet.ExpCloseDate=date()

			Dim fso
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			On Error Resume Next
			Call fso.DeleteFile(Server.MapPath("/baxasr/upload/" & Request("delFile")), True)
			If Err.Number = 0 Then
				'Response.Write "File Deleted"
			Else
				msg ="File is not deleted." & Err.Description & Request("delFile")
			End If
			Set fso = Nothing
		End If '--pageAction = ADD --'
	End If '-- Access Failed --'

	Set mobj = Server.CreateObject("ASRTrans.clsList")
	deptList = mobj.PopulateDept()
	dt= mobj.RetrieveDate
	Set mobj = nothing

%>

<html>
<head>
<Title>ASR - New Request </Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">

</head>
<script language="javascript" src="datepicker.js"></script>
<script language="Javascript" src="fnlist.js"></script>

<script>
var popWin;
function showPopUp(skey) {
	var url = "poplist.asp" + "?skey=" + skey ;
	popWin = open(url, "popupwin", "toolbar=no,scrollbars=yes,width=480,height=300,top=250, left=200");
}

function returnCustAppl(customer,appl,appmgr) {
	frmAdd.customer.value = customer;
	frmAdd.appl.value = appl;
	frmAdd.appmgr.value = appmgr;
}

/*
function returnCustomer(customer) {
	frmAdd.customer.value = customer;
}

function returnAppl(appl) {
	frmAdd.appl.value = appl;
}
*/
//function returnITAccMgr() {

//	if ((!frmAdd.customer.value =="") && (!frmAdd.appl.value ==""))


function returnReqUser(requestor, reqemail, dept, reqmgr, reqmgremail) {
	frmAdd.requestor.value = requestor;
	frmAdd.reqemail.value = reqemail;
	frmAdd.dept.value = dept;
	frmAdd.reqmgr.value = reqmgr;
	frmAdd.reqmgremail.value = reqmgremail;
}
function returnReqUserMgr(reqmgr, reqmgremail) {
	frmAdd.reqmgr.value = reqmgr;
	frmAdd.reqmgremail.value = reqmgremail;
}

//--- ATTACHMENTS SECTION -----
function attach() {
	popWin = open("attachment.asp", "popupwin", "toolbar=no,scrollbars=no,width=430,height=180,top=150, left=170");
}
function returnFile(fname,fnamedis) {
	//alert(fname + "-" + fnamedis);
	document.frmAdd.fname1.value = fname;
	document.frmAdd.fname1dis.value = fnamedis;
	document.frmAdd.pageAction.value = "UPLOAD";
	document.frmAdd.submit();
}

function DeleteRow(row)
	{

	if (confirm("Do you want to delete the file?"))
		{
			document.frmAdd.pageAction.value = "DEL";
			document.frmAdd.delFile.value = document.frmAdd.fname(row+1).value;
			document.frmAdd.delIndex.value = row;
			//alert(document.frmAdd.delFile.value);
			document.frmAdd.submit();
		}

	}

//--SUBMIT FORM---
function submitForm()
{
	var frmobj = document.forms[0];
//alert("msg");

	if(trim(frmobj.requestor.value) == "")
		{
		alert("Please select the requestor name ");
        frmobj.requestor.focus();
        return false;
		}
	if(trim(frmobj.reqemail.value) == "")
		{
		alert("Please enter the requestor email ");
        frmobj.reqemail.focus();
        return false;
		}

	if(trim(frmobj.reqmgr.value) == "")
		{
		alert("Please select requestor's Manager ");
        frmobj.reqmgr.focus();
        return false;
		}
	if(trim(frmobj.reqmgremail.value) == "")
		{
		alert("please enter the manager email ");
        frmobj.reqmgremail.focus();
        return false;
		}

	if (!EmailValidator(frmobj))
		return false;

	if(trim(frmobj.customer.value) == "")
		{
		alert("please select the customer name ");
        frmobj.cmdCustomer.focus();
        return false;
		}
	if(trim(frmobj.appl.value) == "")
		{
		alert("please select the application name ");
        frmobj.cmdAppl.focus();
        return false;
		}

	//if((frmobj.appmgr.value == ""))
	//	{
	//	alert("Accounts manager cannot be blank. Please change the application/customer/dept. ");
    //    frmobj.appl.focus();
    //    return false;
	//	}
	if(trim(frmobj.reqtitle.value) == "")
		{
		alert("please enter the Title ");
        frmobj.reqtitle.focus();
        return false;
		}
	if(trim(frmobj.Desc.value) == "")
		{
		alert("please enter Description ");
        frmobj.Desc.focus();
        return false;
		}
	if(frmobj.Desc.value.length > 250)
	{
		alert("You can enter maximum 250 characters as the Description. To enter more details, Please save the description as text file and attach it.");
		frmobj.Desc.focus();
		return false;
    }
	if((frmobj.expclosedate.value == ""))
		{
		alert("please enter expected Go Live date ");
        frmobj.expclosedate.focus();
        return false;
		}


	document.frmAdd.pageAction.value = "ADD";
	frmobj.submit();

}
</script>
<body >

<% if msg <> "" then %><div class="msginfo"><%= msg %></div><% end if %>
<br>
<form method="post"  id=form1 name=frmAdd>
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="delFile" value="<%=Request("delFile")%>">
<INPUT type="hidden" name="delIndex" value="<%=Request("delIndex")%>">
<INPUT type="hidden" name="UID" value="<%=strUser%> ">

	<table width="550" border="1" cellspacing="0" cellpadding="5" borderColor=salmon>
		<tr>
		<td class="lbl1" width="150" >Action</td>
		<td> <select name="select" style="WIDTH: 165px">
				<option value="1" selected>Save Details</option>
			</select>
			<INPUT id=cmd1 type=button value=OK name=cmd1 style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:submitForm();">
		</td>
		<td width="120"> <INPUT id=cmd2 type=button size=7 value="Attachments" name=cmd2 style="WIDTH: 110px; HEIGHT: 22px" onClick="javascript:attach();">
		</td>
		</tr>
	</table>

	<br><!-- r1 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Request Details</td>
			</tr>
		</table>
	</td></tr>

	<tr><td>
		<table width="550" border="0" cellspacing="2" cellpadding="5">
			<tr>
			<td width="148" class="lbl1">Requestor</td>
		  <td><input name="requestor" id="requestor" maxlength=50 style="WIDTH: 263px; HEIGHT: 20px" value="<%= clsReqDet.ReqUser %>" size=37 >
			  <input name="cmdRequestor" type="button" id="cmdRequestor" value="..." onClick="javascript:showPopUp('USER');">
		  </td>
		</tr>
    <tr>
      <td width="148" class="lbl1">Email</td>
      <td><input name="reqemail" id="reqemail"  maxlength=50 style="WIDTH: 261px; HEIGHT: 20px" size=37 value="<%= clsReqDet.ReqEmail %>" ></td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Requestor's Manager</td>
      <td>
        <input name="reqmgr" id="reqmgr"  maxlength=50  style="WIDTH: 261px; HEIGHT: 20px" value="<%= clsReqDet.ReqUserMgr %>" size=3 >
        <input name="cmdReqUserMgr" type="button" id="cmdReqUserMgr" value="..." onClick="javascript:showPopUp('USERMGR');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Email</td>
      <td><input name="reqmgremail" id="reqmgremail"  maxlength=50 style="WIDTH: 260px; HEIGHT: 20px" size=37 value="<%= clsReqDet.ReqMgrEmail %>"> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Dept</td>
      <td> <select name="dept" id="dept" style="WIDTH: 165px"><%= deptList %></select><script>SetItemValue("dept","<%= clsReqDet.deptname %>");</script>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Customer</td>
      <td>
        <input name="customer" id="customer" type="text" value="<%= clsReqDet.CName %>" readOnly>
        <input name="cmdCustomer" type="button" value="..." onClick="javascript:showPopUp('CUSTAPPL');">
      </td>
    </tr>
	<tr>
      <td width="148" class="lbl1">Application</td>
      <td>
        <input name="appl" id="appl" type="text" value="<%= clsReqDet.ApplName %>" readOnly>
        <input name="cmdAppl" type="button" value="..." onClick="javascript:showPopUp('CUSTAPPL');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">IT Account Manager</td>
      <td><INPUT name="appmgr" id="appmgr"  value="<%= clsReqDet.AppMgr %>" readOnly> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Priority</td>
      <td> <select name="priority" id="priority"
            style="WIDTH: 140px">
          <option value=1>Critical</option>
          <option value=2>Important &amp; Urgent</option>
          <option value=3>Important</option>
          <option value=4>Nice to have</option>
          <option value=5>Suggestion</option>
        </select>
        <script>SetItemValue("priority","<%= clsReqDet.priority %>");</script>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Request Title</td>
      <td> <input name="reqtitle" id="reqtitle"   maxlength=50 style="WIDTH: 318px; HEIGHT: 20px" value="<%= clsReqDet.ReqTitle %>" size=46></td>
    </tr>
    <tr>
      <td class="lbl1" style="VERTICAL-ALIGN: top">Description</td>
      <td><TEXTAREA name="Desc" id=Desc rows=10 cols=50 ><%= clsReqDet.ReqDesc %></TEXTAREA></td>
    </tr>

    <tr>
      <td width="148" class="lbl1">Expected Go Live Date</td>
      <td><INPUT name="expclosedate" id="expclosedate" value=<%= clsReqDet.ExpCloseDateStr %>>
		<a href="javascript:show_calendar('frmAdd.expclosedate');" onMouseOver="window.status='Select Expected Go Live Date';return true;" onMouseOut="window.status='';return true;">
        <img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
      </td>

    </tr>
    <tr>
      <td width="148" class="lbl1">Request Date</td>
      <td> </td>
    </tr>
  </table>

</td></tr>
</table><!-- r1 end-->

<BR><!-- r7 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr>
		<td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr"> Attachments </td>

			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td ><input type="hidden" name="fname" value="">
		<input type="hidden" name="fnamedis" value=""></td>
		<td ><input type="hidden" name="fuser" value=""></td>
		<td ><input type="hidden" name="fdate" value=""></td>
		<INPUT type="hidden" name="fname1" value="">
		<INPUT type="hidden" name="fname1dis" value="">
	</tr>
	<tr>
		<td>


		<table width="550" border="0" cellspacing="2" cellpadding="5">
		<tr>
			<td class="lbl1">DEL</td>
			<td class="lbl1">File Name</td>
			<td class="lbl1">Added By</td>
			<td class="lbl1">Added Date</td>
		</tr>

<%

		arrcount = Request.Form("fname").Count
		introw=0

		if pageAction = "DEL" then
			for i = 2 to arrcount
				if cint(i-2) <> cint(delIndex) then

%>
		<tr>
			<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=introw%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>
			<td style="WIDTH: 120px" width=120><a href="../upload/<%= Request.Form("fname")(i) %>"><%= Request.Form("fnamedis")(i) %></a>
				<input type="hidden" name="fname" value="<%= Request.Form("fname")(i) %>">
				<input type="hidden" name="fnamedis" value="<%= Request.Form("fnamedis")(i) %>">
				</td>
			<td style="WIDTH: 120px" width=120><%= Request.Form("fuser")(i) %><input type="hidden" name="fuser" value="<%= Request.Form("fuser")(i) %>"></td>
			<td style="WIDTH: 120px" width=120><%= Request.Form("fdate")(i) %><input type="hidden" name="fdate" value="<%= Request.Form("fdate")(i) %>"></td>
		</tr>
<%
				introw=introw+1
				end if

			next
		Else
			for i = 2 to arrcount

%>
		<tr>
			<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=introw%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>
			<td style="WIDTH: 120px" width=120><a href="../upload/<%= Request.Form("fname")(i) %>"><%= Request.Form("fnamedis")(i) %></a>
				<input type="hidden" name="fname" value="<%= Request.Form("fname")(i) %>">
				<input type="hidden" name="fnamedis" value="<%= Request.Form("fnamedis")(i) %>">
				</td>
			<td style="WIDTH: 120px" width=120><%= Request.Form("fuser")(i) %><input type="hidden" name="fuser" value="<%= Request.Form("fuser")(i) %>"></td>
			<td style="WIDTH: 120px" width=120><%= Request.Form("fdate")(i) %><input type="hidden" name="fdate" value="<%= Request.Form("fdate")(i) %>"></td>
		</tr>
<%
			introw=introw+1

			next
		end if
		if (pageAction = "UPLOAD") then
%>
		<tr>
			<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=introw%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>
			<td style="WIDTH: 120px" width=120><a href="../upload/<%= Request.Form("fname1") %>"><%= Request.Form("fname1dis") %></a>
				<input type="hidden" name="fname" value="<%= Request.Form("fname1") %>">
				<input type="hidden" name="fnamedis" value="<%= Request.Form("fname1dis") %>">
				</td>
			<td style="WIDTH: 120px" width=120><%= strUser  %>
				<input type="hidden" name="fuser" value="<%= strUser %>"></td>
			<td style="WIDTH: 120px" width=120><%= now() %>
				<input type="hidden" name="fdate" value="<%=now() %>"></td>
		</tr>
<%
		end if
%>
		</table>
		</td>
	</tr>
</table>
<br>
<table width="550" border="1" cellspacing="0" cellpadding="5" borderColor=salmon>
		<tr>

		<td align="right"> <INPUT id=cmd2 type=button size=7 value="Attachments" name=cmdAttach style="WIDTH: 110px; HEIGHT: 22px" onClick="javascript:attach();">
		</td>
		</tr>
	</table>
</TD></TR></TABLE><!-- r7 end-->
</form>
</body>
</html>
