<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option Explicit%>

<%
	Dim objReqDet, strUser, retVal, mobj, objReqCost
	Dim rs, deptList, msg, hrRate, dt, introw, rsReqCost
	Dim i, arrcount, pageAction, delFile, delIndex

	strUser=Request.ServerVariables ("LOGON_USER")
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)
	'strUser = "hema"

	Set objReqDet = Server.CreateObject("ASRTrans.clsRequestDets")
	Set objReqCost = Server.CreateObject("ASRTrans.clsReqCost")
	objReqCost.ReqNo = Request("ReqNo")
	
	If strUser="" then
		msg = "Access Denied"
	Else
		pageAction = Request.Form("pageAction")
		delFile = Request.Form("delFile")
		delIndex = Request.Form("delIndex")
		If pageAction  = "" then
			objReqDet.ReqNo = Request("ReqNo")
			'Response.Write objReqDet.ReqNo
			On Error Resume Next
			objReqDet.RetrieveRecord
			If Err Then
				msg = "There is some error in getting your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			End If
		ElseIf pageAction = "UPD" then
			On Error Resume Next
			objReqDet.SetValuesUpd Request
			If Err Then
				msg = "There is some error in saving your details(prepare data) please try again! <br>Error " & Err.Number & " : " & Err.Description
			end if
			
			Dim fso
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			On Error Resume Next
			For i = 2 to arrcount
				Call fso.DeleteFile(Server.MapPath("../upload/" & delFile(i)), True)
			Next

			If Err.Number = 0 Then
				'Response.Write "File Deleted"
			Else
				msg ="File is not deleted." & Err.Description
			End If
			Set fso = Nothing

			if msg ="" then
				retVal = objReqDet.UpdateRecord
				'Response.Write objReqDet.RetrieveStatusNo(objReqDet.SCode)
				If Err Then
					msg = "There is some error in saving your details(upd) please try again! <br>Error " & Err.Number & " : " & Err.Description
				Else
					If retVal = "OK" then
					'Response.Redirect "search.asp?pageAction=SEARCH"
					%>

						<script>
						    var parentWin = opener;

						    try {
						        if (opener.document.title == "ASR - Search") {
						            opener.document.frmSearch.pageAction.value = "SEARCH";
						            opener.document.frmSearch.reqno.value = '<%=Request("ReqNo")%>';
						            opener.document.frmSearch.submit();
						        }
						        window.close();

						    }
						    catch (e) {
						        alert("Your details has been updated. The destination (ASR - SEARCH) window has been closed. Please go to the ASR Search screen to modify the request details");
						        document.location = "search.asp"
						    }



						</script>
					<%
					Else
							msg = "There is some error in saving your details please try again!"
					End If
				End if
			End If
		ElseIf pageAction = "CANCEL" then
			On Error Resume Next
			retVal = objReqDet.UpdateStatusCH(Request)
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then
					'Response.Redirect "search.asp?pageAction=SEARCH"
					%>

					<script>
					    var parentWin = opener;

					    try {
					        if (opener.document.title == "ASR - Search") {
					            opener.document.frmSearch.pageAction.value = "SEARCH";
					            opener.document.frmSearch.reqno.value = '<%=Request("ReqNo")%>';
					            opener.document.frmSearch.submit();
					        }

					    }
					    catch (e) {
					        alert("The destination (ASR - SEARCH) window has been closed. Please go to the ASR Search screen to modify the request details");
					    }
					    window.close();
					</script>
					<%
				Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "HOLD" then
			On Error Resume Next
			retVal = objReqDet.UpdateStatusCH(Request)
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then
						'Response.Redirect "search.asp?pageAction=SEARCH"
					%>

					<script>
					    var parentWin = opener;

					    try {
					        if (opener.document.title == "ASR - Search") {
					            opener.document.frmSearch.pageAction.value = "SEARCH";
					            opener.document.frmSearch.reqno.value = '<%=Request("ReqNo")%>';
					            opener.document.frmSearch.submit();
					        }

					    }
					    catch (e) {
					        alert("The destination (ASR - SEARCH) window has been closed. Please go to the ASR Search screen to modify the request details");
					    }
					    window.close();
					</script>
				<%
				Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If
		ElseIf pageAction = "RELEASE" then
			On Error Resume Next
			objReqDet.ReqNo = Request("ReqNo")
			objReqDet.ModUser = strUser
			retVal = objReqDet.UpdateRelease
			If Err Then
				msg = "There is some error in saving your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			Else
				If retVal = "OK" then
						'Response.Redirect "search.asp?pageAction=SEARCH"
						'
					%>

					<script>
					    var parentWin = opener;

					    try {
					        if (opener.document.title == "ASR - Search") {
					            opener.document.frmSearch.pageAction.value = "SEARCH";
					            opener.document.frmSearch.reqno.value = '<%=Request("ReqNo")%>';
					            opener.document.frmSearch.submit();
					        }

					    }
					    catch (e) {
					        alert("The destination (ASR - SEARCH) window has been closed. Please go to the ASR Search screen to modify the request details");
					    }
					    window.close();
					</script>
				<%
				Else
						msg = "There is some error in saving your details please try again!"
				End If
			End If

		ElseIf pageAction = "UPLOAD" then
			objReqDet.SetValuesUpd Request

		ElseIf pageAction = "DEL" then
			'clsReqDet.ExpCloseDate=date()
			'
			objReqDet.SetValuesUpd Request
			
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			On Error Resume Next
			Call fso.DeleteFile(Server.MapPath("../upload/" & Request("delFile")), True)
			If Err.Number = 0 Then
				'Response.Write "File Deleted"
			Else
				msg ="File is not deleted." & Err.Description & Request("delFile")
			End If
			Set fso = Nothing
		ElseIf pageAction = "UNLOAD" then
			Set objReqDet = nothing
			Set mobj = nothing
		End If '--pageAction =--'
	End If '-- Access Failed --'

	Set mobj = Server.CreateObject("ASRTrans.clsList")
	deptList = mobj.PopulateDept()
	dt = mobj.RetrieveDate()
	hrRate = mobj.RetrieveHourRate()

%>

<html>
<head>
<Title>ASR - Edit Request Details</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<!-- #include file="access2.asp" -->

<script language="javascript" src="validate.js"></script>
<script language="javascript" src="datepicker.js"></script>
<script language="Javascript" src="fnlist.js"></script>
<script language="Javascript" src="workflow.js"></script>
<script language="javascript" src="status.js"></script>

<script>
    var popWin;
    function showPopUp(skey) {
        var url = "poplist.asp" + "?skey=" + skey;
        popWin = open(url, "popupwin", "toolbar=no,scrollbars=yes,width=480,height=300,top=250, left=200");
    }
    function returnCustAppl(customer, appl, appmgr) {
        frmEdit.customer.value = customer;
        frmEdit.appl.value = appl;
        frmEdit.appmgr.value = appmgr;
    }
    function returnCustomer(customer, appmgr) {
        frmEdit.customer.value = customer;
        frmEdit.appmgr.value = appmgr;
    }
    function returnAppl(appl, appmgr) {
        frmEdit.appl.value = appl;
        frmEdit.appmgr.value = appmgr;
    }
    function returnReqUser(requestor, reqemail, dept, reqmgr, reqmgremail) {
        frmEdit.requestor.value = requestor;
        frmEdit.reqemail.value = reqemail;
        frmEdit.dept.value = dept;
        frmEdit.reqmgr.value = reqmgr;
        frmEdit.reqmgremail.value = reqmgremail;
    }
    function returnReqUserMgr(reqmgr, reqmgremail) {
        frmEdit.reqmgr.value = reqmgr;
        frmEdit.reqmgremail.value = reqmgremail;
    }
    function returnTeamLead(TeamLead) {
        frmEdit.TeamLead.value = TeamLead;
    }
    function returnDeveloper(Developer) {
        frmEdit.Developer.value = Developer;
    }
    function returnUserUAT(UserUAT) {
        frmEdit.UserUAT.value = UserUAT;
    }
    function returnActQC(ActQC) {
        frmEdit.ActQC.value = ActQC;
    }
    function returnExpQC(ExpQC) {
        frmEdit.ExpQC.value = ExpQC;
    }
    function returnDeployUser(deployuser) {
        frmEdit.DeployUser.value = deployuser;
    }
    function GetCostExp(hrRate) {
        frmEdit.EstTotalCost.value = frmEdit.EstManHour.value * hrRate;
    }
    function GetCostAct(hrRate) {
        frmEdit.ActTotalCost.value = frmEdit.ActManHour.value * hrRate;
    }

    //--- ATTACHMENTS SECTION -----
    function attach() {
        popWin = open("attachment.asp", "popupwin", "toolbar=no,scrollbars=no,width=430,height=150,top=150, left=170");
    }
    function returnFile(fname, fnamedis) {
        //alert("Attachment : " + fname + " accepted!");
        document.frmEdit.fname1.value = fname;
        document.frmEdit.fname1dis.value = fnamedis;
        document.frmEdit.pageAction.value = "UPLOAD";
        document.frmEdit.submit();
    }
    function DeleteRow(row) {

        if (confirm("Do you want to delete the file?")) {
            document.frmEdit.pageAction.value = "DEL";
            document.frmEdit.delFile.value = document.frmEdit.fname(row).value;
            document.frmEdit.delIndex.value = row;
            //alert(document.frmEdit.delFile.value);
            document.frmEdit.submit();
        }

    }

    //--SUBMIT FORM---
    function submitForm(action) {
        var msg = "";
        var tmp = "";
        var ret = "";
        var frmobj = document.forms[0];
        //alert(document.frmEdit.action.value);
        //alert for empty related required fields
        if (document.frmEdit.action.value == 1) {

            if (!formRequestDetails(document.frmEdit))
                return false;

            if (!EmailValidator(document.frmEdit))
                return false;

            if (!formITQCDetails(document.frmEdit))
                return false;

            if (!formITDetails(document.frmEdit))
                return false;

            if (!formProgStart(document.frmEdit))
                return false;

            if (!formProgDetails(document.frmEdit))
                return false;
            if ((frmEdit.StatusQC.value == "P") && (frmEdit.SCode.value != "UAT") && (frmEdit.SCode.value != "Deploy")) {
                //frmEdit.UATReadyDate.value ='<%=date()%>';
            }

            if (!formQCDetails(document.frmEdit))
                return false;

            if (!StatusCheck(document.frmEdit))
                return false;
            if (!formUatDetails(document.frmEdit))
                return false;

            if (!formDeployDetails(document.frmEdit))
                return false;

            if (!formProgBlank(document.frmEdit))
                return false;

            if (!formQCBlank(document.frmEdit))
                return false;

            if (!formUATBlank(document.frmEdit))
                return false;

            if (!formDeployBlank(document.frmEdit))
                return false;

            //End of checking the related fields

            //check remark cols
            if (frmEdit.Desc.value.length > 250) {
                alert("The Description column exceeded the maximum limit (250 characters). ");
                frmEdit.Desc.focus();
                return false;
            }
            if (frmEdit.RemarksIT.value.length > 250) {
                alert("The IT Remarks column exceeded the maximum limit (250 characters). ");
                frmEdit.RemarksIT.focus();
                return false;
            }
            if (frmEdit.RemarksDev.value.length > 250) {
                alert("The Progress Details column exceeded the maximum limit (250 characters). ");
                frmEdit.RemarksDev.focus();
                return false;
            }
            if (frmEdit.RemarksQC.value.length > 250) {
                alert("The QC Remarks column exceeded the maximum limit (250 characters). ");
                frmEdit.RemarksQC.focus();
                return false;
            }
            if (frmEdit.RemarksUAT.value.length > 250) {
                alert("The UAT Remarks column exceeded the maximum limit (250 characters). ");
                frmEdit.RemarksUAT.focus();
                return false;
            }
            if (frmEdit.RemarksDeploy.value.length > 250) {
                alert("The Deploy Remarks column exceeded the maximum limit (250 characters). ");
                frmEdit.RemarksDeploy.focus();
                return false;
            }

            //check date values
            if (isDate(frmEdit.expclosedate.value, 'dd/MM/yyyy') == false) {
                alert("Enter valid Date format (DD/MM/YYYY) ");
                frmEdit.expclosedate.focus();
                return false;
            }

            if (frmEdit.ExpStartDate.value != "") {
                if (isDate(frmEdit.ExpStartDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.ExpStartDate.focus();
                    return false;
                }

                if (isDate(frmEdit.ExpEndDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.ExpEndDate.focus();
                    return false;
                }
                if (compareDates(frmEdit.ExpStartDate.value, 'dd/MM/yyyy', frmEdit.ExpEndDate.value, 'dd/MM/yyyy') == 1) {
                    alert("Exp. Start Date can not be greater than Exp. Complete date ");
                    frmEdit.ExpEndDate.focus();
                    return false;
                }
            }

            if (frmEdit.ActStartDate.value != "") {
                if (isDate(frmEdit.ActStartDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.ActStartDate.focus();
                    return false;
                }
                if (compareDates(frmEdit.ActStartDate.value, 'dd/MM/yyyy', frmEdit.dtval.value, 'dd/MM/yyyy') == 1) {
                    alert("You cannot enter any future date as Actual Start Date.");
                    frmEdit.ActStartDate.focus();
                    return false;
                }
            }
            if (frmEdit.ActEndDate.value != "") {
                if (isDate(frmEdit.ActEndDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.ActEndDate.focus();
                    return false;
                }
                if (compareDates(frmEdit.ActStartDate.value, 'dd/MM/yyyy', frmEdit.ActEndDate.value, 'dd/MM/yyyy') == 1) {
                    alert("Act. Start Date can not be greater than Act. Completed date ");
                    frmEdit.ActEndDate.focus();
                    return false;
                }
                if (compareDates(frmEdit.ActEndDate.value, 'dd/MM/yyyy', frmEdit.dtval.value, 'dd/MM/yyyy') == 1) {
                    alert("You cannot enter any future date as Actual Completed Date.");
                    frmEdit.ActEndDate.focus();
                    return false;
                }
            }

            if ((frmEdit.ActQC.value == frmEdit.Developer.value) && (frmEdit.Developer.value != "")) {
                alert("The Actual QC cannot be the Developer. Please select different name.");
                frmEdit.ActQC.focus();
                return false;
            }

            if ((frmEdit.StatusQC.value == "F") && (frmEdit.SCode.value == "UAT")) {
                alert("The request status is already chagned to UAT. You cannot modify the QC status now.");
                frmEdit.StatusQC.focus();
                return false;
            }

            if (frmEdit.StatusQC.value == "F") {
                if (frmEdit.UATReadyDate.value != "") {
                    alert("You cannot enter UAT Ready Date for QC Status:Failed ");
                    frmEdit.UATReadyDate.focus();
                    return false;
                }
            }
            if (frmEdit.UATReadyDate.value != "") {
                if (isDate(frmEdit.UATReadyDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.UATReadyDate.focus();
                    return false;
                }
                if (frmEdit.StatusQC.value == "P") {
                    if (compareDates(frmEdit.ActEndDate.value, 'dd/MM/yyyy', frmEdit.UATReadyDate.value, 'dd/MM/yyyy') == 1) {
                        alert("Act. Completed Date can not be greater than UAT Ready Date ");
                        frmEdit.UATReadyDate.focus();
                        return false;
                    }
                }
            }
            if ((frmEdit.requestor.value != frmEdit.UserUAT.value) && (frmEdit.UserUAT.value != "")) {
                ret = confirm("The Requestor and the UAT user are different. The Requestor should be the right person to conduct the UAT. Do you want to continue?");
                if (ret) {
                }
                else {
                    frmEdit.UserUAT.focus();
                    return false;
                }
            }

            if ((frmEdit.StatusUAT.value == "F") && (frmEdit.SCode.value == "Deploy")) {
                alert("The request status is already chagned to 'Deploy'. You cannot modify the UAT status now.");
                frmEdit.StatusUAT.focus();
                return false;
            }

            if (frmEdit.StatusUAT.value == "F") {
                if (frmEdit.ExpCutinDate.value != "") {
                    alert("You cannot enter Exp. Cut in Date for UAT Status:Failed ");
                    frmEdit.ExpCutinDate.focus();
                    return false;
                }
            }
            if ((frmEdit.StatusUAT.value == "F") && (frmEdit.SCode.value != "UAT") && (frmEdit.UATFailed.value == " ")) {
                alert("Please clear the UAT details and save the Request details. The Request status must be UAT when you set the UAT failed status. ");
                frmEdit.StatusUAT.focus();
                return false;
            }


            if ((frmEdit.StatusUAT.value == "F") && (frmEdit.SCode.value == "UAT")) {
                frmEdit.UATFailed.value = "Y"
            }

            if (frmEdit.ExpCutinDate.value != "") {
                if (isDate(frmEdit.ExpCutinDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.ExpCutinDate.focus();
                    return false;
                }
                if (compareDates(frmEdit.UATReadyDate.value, 'dd/MM/yyyy', frmEdit.ExpCutinDate.value, 'dd/MM/yyyy') == 1) {
                    alert("UAT Ready Date can not be greater than Exp. Cut in Date ");
                    frmEdit.ExpCutinDate.focus();
                    return false;
                }
            }
            if (frmEdit.ActCutinDate.value != "") {
                if (isDate(frmEdit.ActCutinDate.value, 'dd/MM/yyyy') == false) {
                    alert("Enter valid Date format (DD/MM/YYYY) ");
                    frmEdit.ActCutinDate.focus();
                    return false;
                }
                //	if (compareDates(frmEdit.ActCutinDate.value,'dd/MM/yyyy',frmEdit.ExpCutinDate.value,'dd/MM/yyyy')==1)
                //		{
                //		alert("Act. Cut in Date can not be greater than Exp. Cut in Date ");
                //		frmEdit.ActCutinDate.focus();
                //		return false;
                //		}
                if (compareDates(frmEdit.UATReadyDate.value, 'dd/MM/yyyy', frmEdit.ActCutinDate.value, 'dd/MM/yyyy') == 1) {
                    alert("Act. Cut in Date can not be less than UAT Ready Date ");
                    frmEdit.ActCutinDate.focus();
                    return false;
                }
                if (compareDates(frmEdit.ActCutinDate.value, 'dd/MM/yyyy', frmEdit.dtval.value, 'dd/MM/yyyy') == 1) {
                    alert("You cannot enter any future date for Actual Cut-in date.");
                    frmEdit.ActCutinDate.focus();
                    return false;
                }

            }
            if ((frmEdit.StatusQC.value == "P") && (frmEdit.SCode.value != "UAT") && (frmEdit.SCode.value != "Deploy")) {
                ret = confirm("Have you set up the UAT environment? Do you want to continue?");
                if (ret) {
                }
                else {
                    return false;
                }
            }

            //---  No more required 
            //---if ((frmEdit.ProjCost.value =="NO") && (trim(frmEdit.DeployUser.value) !="") && (frmEdit.reqtype.value =="P"))
            //---{
            //---	alert("The project cost entries must be entered. You cannot close the request now.");
            //---	frmEdit.RemarksDeploy.focus();
            //---	return false;
            //--- }

            document.frmEdit.pageAction.value = "UPD";
        }

        if (document.frmEdit.action.value == 2) {
            ret = confirm("Do you want to Cancel this request");
            if (ret) {
                document.frmEdit.pageAction.value = "CANCEL";
                document.frmEdit.StatusCH.value = "Cancelled";
                tmp = prompt("Enter the Reason for the Cancel:", "", "ASR - Cancel");
                if (!tmp == "") {
                    if (tmp.length > 250) {
                        alert("The reason exceeded the maximum limit (250 characters). ");
                        return false;
                    }
                    document.frmEdit.RemarksCH.value = tmp;
                }
                else {
                    alert("Please enter the reason for the Cancel");
                    return false;
                }
            }
            else {
                return false;
            }
        }
        if (document.frmEdit.action.value == 3) {
            ret = confirm("Do you want to Hold this request");
            if (ret) {
                document.frmEdit.pageAction.value = "HOLD";
                document.frmEdit.StatusCH.value = "Hold";
                tmp = prompt("Enter the Reason for the Hold :", "", "ASR - Hold");
                if (!tmp == "") {
                    if (tmp.length > 250) {
                        alert("The reason exceeded the maximum limit (250 characters). ");
                        return false;
                    }

                    document.frmEdit.RemarksCH.value = tmp;
                }
                else {
                    alert("Please enter the reason for the Hold");
                    return false;
                }
            }
            else {
                return false;
            }
        }
        if (document.frmEdit.action.value == 4) {
            ret = confirm("Do you want to Release this request");
            if (ret) {
                document.frmEdit.pageAction.value = "RELEASE";
            }
            else {
                return false;
            }
        }
        frmobj.submit();
    }

    Date.prototype.Format = function (fmt) {
        var o = {
            "M+": this.getMonth() + 1,
            "d+": this.getDate(),
            "h+": this.getHours(),
            "m+": this.getMinutes(),
            "s+": this.getSeconds(),
            "q+": Math.floor((this.getMonth() + 3) / 3),
            "S": this.getMilliseconds()
        };
        if (/(y+)/.test(fmt))
            fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
        for (var k in o)
            if (new RegExp("(" + k + ")").test(fmt))
                fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
        return fmt;
    };

    function getValue() {
        //alert(1);
        document.frmEdit.dtval.value = new Date().Format("dd/MM/yyyy");

    }




</script>

<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<body onload=getValue()>
<form method="post"  id=form1 name=frmEdit action="editrequest.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="StatusCH" value="">
<INPUT type="hidden" name="ReqNo" value="<%=Request("ReqNo")%>">
<INPUT type="hidden" name="SCode" value="<%=trim(objReqDet.SCode)%>">
<INPUT type="hidden" name="delFile" value="<%=Request("delFile")%>">
<INPUT type="hidden" name="delIndex" value="<%=Request("delIndex")%>">
<INPUT type="hidden" name="UID" value="<%=strUser%>">

<INPUT type="hidden" name="dtval" value="">

<table width="552" border="1" cellspacing="0" cellpadding="3" bordercolor="#9966cc" style="WIDTH: 552px; HEIGHT: 53px">
      <tr vAlign=center Align=center>
      <td class="trHdr" width=80>Request No</td>
      <td class="trHdr" width=80>Current Status</td>
      <td class="trHdr" width=120>Last Modified By</td>
      <td class="trHdr" width=220>Last Modified Date</td>
      </tr>
      <tr>
      <td width=80> <%=Request("ReqNo")%> </td>
      <td width=120> <%=objReqDet.SCode%> </td>
      <% if  objReqDet.ModUser="" then %>
		<td width=120> <%=objReqDet.CrUser%></td>
		<td width=220> <%=objReqDet.CrDate%></td>
      <% Else  %>
		<td width=120> <%=objReqDet.ModUser%></td>
		<td width=220> <%=objReqDet.ModDate%></td>
      <% End if %>
      </tr>
	</table>


	<%if Trim(objReqDet.SCode)="Cancelled" or Trim(objReqDet.SCode)="Closed" then %>

	<%elseif Trim(objReqDet.SCode)="Hold" then %>
		<br>
		<table width="550" border="1" cellspacing="0" cellpadding="5" borderColor=salmon>
			<tr>
			<td class="lbl1" width="150" >Action</td>
			<td> <select name="action" style="WIDTH: 165px">
					<option value="4">Release Request</option>
				</select>
				<INPUT id=cmdOK type=button value=OK name=cmdOK style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:submitForm();">
			</td>
			</tr>
		</table>
	<% Else  %>
		<br>
		<table width="550" border="1" cellspacing="0" cellpadding="5" borderColor=salmon>
			<tr>
			<td class="lbl1" width="150" >Action</td>
			<td> <select name="action" style="WIDTH: 165px">
					<option value="1" selected>Save Details</option>
					<option value="2">Cancel Request </option>
					<option value="3">Hold Request</option>
				</select>
				<INPUT id=cmdOK type=button value=OK name=cmdOK style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:submitForm();">
			</td>
			<td width="120"> <INPUT name=cmdAttach id=cmdAttach type=button size=7 value="Attachments"  style="WIDTH: 110px; HEIGHT: 22px" onClick="javascript:attach();">
			</td>
			</tr>
		</table>
	<% End if %>
	<BR>	
    <BR>
<!-- r1 start-->
	<table width="555" border="1" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Request/Project Details</td>
			</tr>
		</table>
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    
	
	<tr><td>
	<table width="550" border="0" cellspacing="2" cellpadding="5">
	<tr>
      <td width="148" class="lbl1">Request Type</td>
      <td> <select name="reqtype" id="reqtype" disabled
            style="WIDTH: 140px">
          <option value="R">Request</option>
          <option value="P">Project</option>
        </select>
        <script>            SetItemValue("reqtype", "<%= objReqDet.reqtype %>");</script>
      </td>
    </tr>
		<tr>
			<td width="148" class="lbl1">Requestor</td>
		  <td><input name="requestor" style="WIDTH: 263px; HEIGHT: 20px" maxlength=50 value="<%=objReqDet.ReqUser%>" size=37 >
		  <input name="cmdRequestor" type="button" id="cmdRequestor" value="..." onClick="javascript:showPopUp('USER');">
		  </td>
		</tr>
    <tr>
      <td width="148" class="lbl1">Email</td>
      <td><input name="reqemail" style="WIDTH: 261px; HEIGHT: 20px" maxlength=50 size=37 value="<%=objReqDet.ReqEmail%>" ></td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Requestor's Manager</td>
      <td>
        <input name="reqmgr" style="WIDTH: 261px; HEIGHT: 20px" maxlength=50 value="<%=objReqDet.ReqUserMgr%>" size=3 >
        <input name="cmdReqUserMgr" type="button" id="cmdReqUserMgr" value="..." onClick="javascript:showPopUp('USERMGR');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Email</td>
      <td><input name="reqmgremail" style="WIDTH: 260px; HEIGHT: 20px" maxlength=50 size=37 value="<%=objReqDet.ReqMgrEmail%>"> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Dept</td>
      <td> <select name="dept" style="WIDTH: 165px"><%=deptList%></select><script>                                                                              SetItemValue("dept", "<%=objReqDet.deptname%>");</script>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Customer</td>
      <td >
        <input name="customer"  type="text" value="<%=objReqDet.CName%>" readOnly> <input name="cmdCustomer" type="button" value="..." onClick="javascript:showPopUp('CUSTAPPL');">
      </td>
    </tr>
	<tr>
      <td width="148" class="lbl1">Application</td>
      <td>
        <input name="appl" type="text" value="<%=objReqDet.ApplName%>" readOnly> <input name="cmdAppl" type="button" value="..." onClick="javascript:showPopUp('CUSTAPPL');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">IT Account Manager</td>
      <td><INPUT name=appmgr style="WIDTH: 260px; HEIGHT: 20px" value="<%=objReqDet.AppMgr%>" readOnly></td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Priority</td>
      <td> <select name="priority"
            style="WIDTH: 140px">
          <option value="1">Critical</option>
          <option value="2">Important &amp; Urgent</option>
          <option value="3">Important</option>
          <option value="4">Nice to have</option>
          <option value="5">Suggestion</option>
        </select>
        <script>            SetItemValue("priority", "<%=objReqDet.Priority%>");</script>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Request Title</td>
      <td> <input name="reqtitle" style="WIDTH: 318px; HEIGHT: 20px" maxlength=50 value="<%=objReqDet.ReqTitle%>" size=46></td>
    </tr>
    <tr>
      <td class="lbl1" style="VERTICAL-ALIGN: top">Description</td>
    <td><TEXTAREA id=Desc name=Desc rows=10 cols=50 ><%=objReqDet.ReqDesc%></TEXTAREA></td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Expected Go Live Date</td>
      <td><INPUT name=expclosedate value="<%=objReqDet.ExpCloseDatestr%>">
	      <a href="javascript:show_calendar('frmEdit.expclosedate');" onMouseOver="window.status='Select Expected Close Date';return true;" onMouseOut="window.status='';return true;">
		  <img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>

      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Request Date</td>
      <td><%=objReqDet.CrDate%> </td>
    </tr>
    </table>

	</td></tr>
  </table><!-- r1 end-->
<BR>
<!-- rem start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
          <tr><td class="trHdr"> Cancel/Hold Remarks </td></tr>
		</table>
		</td>
	</tr>
	<tr><td>
		<table width="550" border="0" cellspacing="2" cellpadding="5">
		<tr>
			<td><TEXTAREA name=RemarksCH id=RemarksCH style="WIDTH: 526px; HEIGHT: 49px"  cols=79 readonly> <%=objReqDet.RemarksCH%> </TEXTAREA></td>
		</tr>
		</table>
    </td></tr>
  </table><!-- rem end-->

<BR><!-- r2 start-->

  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr><td class="trHdr">IT Details</td>
		    </tr>
		</table>
	</td></tr>
	<tr><td>
		<table width="550" border="0" cellspacing="2" cellpadding="5">
			<tr>
				<td width="148" class="lbl1">Team Leader</td>
				<td><input name="TeamLead" value="<%=objReqDet.TeamLead%>" readonly style="WIDTH: 263px; HEIGHT: 20px" size=37>
				    <input name="cmdTeamLead" type="button" id="cmdTeamLead" value="..." onClick="javascript:showPopUp('TeamLead');">
				</td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Team Members</td>
				<td> <input name=Developer value="<%=objReqDet.Developer%>" readonly maxlength=100  style="WIDTH: 263px; HEIGHT: 20px" size=37>
					 <input name="cmdDeveloper" type="button" id="cmdDeveloper" value="..." onClick="javascript:showPopUp('Developer');">
				</td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Quality Control</td>
				<td><input name="ExpQC" value="<%=objReqDet.ExpQC%>" readonly style="WIDTH: 263px; HEIGHT: 20px" size=37>
				    <input name="cmdExpQC" type="button" id="cmdExpQC" value="..." onClick="javascript:showPopUp('ExpQC');">
				</td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Estimated ManHour</td>
				<td><input name="EstManHour" value="<%=objReqDet.EstManHour%>" maxlength=5 style="WIDTH: 116px; HEIGHT: 20px" size=16
					maxlength="4" onKeyPress="if(!((window.event.keyCode >= '48')&&(window.event.keyCode <= '57'))){alert('Please enter NUMERIC values only.');return false;}"
					onblur="javascript:GetCostExp('<%=objReqDet.EstHourRate%>');"> hrs
				</td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Estimated Cost/Hour</td>
					<td>  <INPUT type=hidden name=EstHourRate value="<%=objReqDet.EstHourRate %> ">
					<%=objReqDet.EstHourRate%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Estimated Total Cost</td>
				<td><INPUT name=EstTotalCost value="<%=objReqDet.EstTotalCost%>" maxlength=8 style="WIDTH: 118px; HEIGHT: 20px" size=15> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Exp. Start Date</td>
				<td><INPUT name=ExpStartDate value="<%=objReqDet.ExpStartDatestr%>" id=ExpStartDate style="WIDTH: 115px; HEIGHT: 20px" size=16 >&nbsp;
						<a href="javascript:show_calendar('frmEdit.ExpStartDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
						<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
				</td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Exp. Complete Date</td>
				<td><INPUT name=ExpEndDate value="<%=objReqDet.ExpEndDatestr%>" id=ExpEndDate style="WIDTH: 115px; HEIGHT: 20px" size=16 >&nbsp;
						<a href="javascript:show_calendar('frmEdit.ExpEndDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
						<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
				</td>
			</tr>
			<tr>
				<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
				<td><TEXTAREA name=RemarksIT id=RemarksIT style="WIDTH: 323px; HEIGHT: 77px"  rows=5 cols=50><%=objReqDet.RemarksIT%></TEXTAREA> </td>
			</tr>
	</table>
	</td></tr>
</table><!-- r2 end-->
<BR>
<%If appUserType="I" or appAccessLevel="1" Then %>
<!-- cost start-->
	<table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td class="trHdr"> Project Cost</td>
	</td></tr>
	<tr><td>
	<table width="550" border="1" cellspacing="1" cellpadding="2">
		
		<% 
		set rsReqCost = objReqCost.RetrieveRequestCostList() 
		

		if rsReqCost.EOF = False then%>
			<INPUT type="hidden" name="ProjCost" value="OK">
			<tr> 
			 <th class="lbl1">&nbsp;Year&nbsp;</th>
			 <th class="lbl1">&nbsp;Month&nbsp;</th>
			 <th class="lbl1">&nbsp;ChargeCost&nbsp;</th>
			 <th class="lbl1">&nbsp;NonChargeCost&nbsp;</th>
			 <th class="lbl1">&nbsp;CrUser&nbsp;</th>
			 <th class="lbl1">&nbsp;CrDate&nbsp;</th>
			</tr>
	<%  
		'SHOW RS CONTENTS'
		'ReqNo, RCYear, RCMonth, ChargeCost, NonChargeCost
		Do While rsReqCost.EOF = False
		'if rsfiles.EOF = False then
%>
		<tr>
			<td style="WIDTH: 120px" width=120 align="center"><%=rsReqCost(1)%></td>
			<td style="WIDTH: 120px" width=120 align="center"><%=rsReqCost(2)%></td>
			<td style="WIDTH: 120px" width=120 align="right"><%=rsReqCost(3)%>&nbsp;</td>
			<td style="WIDTH: 120px" width=120 align="right"><%=rsReqCost(4)%>&nbsp;</td>
			<td style="WIDTH: 120px" width=120><%=rsReqCost(5)%></td>
			<td style="WIDTH: 120px" width=120><%=rsReqCost(6)%></td>
		</tr>
		
		<BR>
<%

		rsReqCost.MoveNext
		Loop
		Else%>
		<INPUT type="hidden" name="ProjCost" value="NO">
		<%End if
		rsReqCost.close
		set rsReqCost = Nothing
		
%>
</table>
</td></tr>
</table>
<BR>
<!-- cost end-->
<%End If  %>
<!-- r3 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
  <table width="550" border="0" cellspacing="0" cellpadding="5">
          <tr>
      <td class="trHdr">Programmer's Update</td>

      <td align="right" class="trHdr" style="WIDTH: 80px">&nbsp;&nbsp;
      </td>
    </tr>
	</table>
	</td>
	</tr>
	<tr>
	<td>

  <table width="550" border="0" cellspacing="2" cellpadding="5">
  </table>
	<table width="550" border="0" cellspacing="2" cellpadding="5">
    <tr>
      <td width="148" class="lbl1" >Actual Start Date</td>
      <td><INPUT name=ActStartDate value="<%=objReqDet.ActStartDateStr%>" id=ActStartDate> &nbsp;

			<a href="javascript:show_calendar('frmEdit.ActStartDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
						<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
      </td>
    </tr>
    </table>
    <HR>
    <table width="550" border="0" cellspacing="2" cellpadding="5">
    <tr>
      <td width="148" class="lbl1" style="VERTICAL-ALIGN: top">Progress Details</td>
      <td><TEXTAREA name=RemarksDev id=RemarksDev style="WIDTH: 323px; HEIGHT: 77px"  rows=5 cols=50><%=objReqDet.RemarksDev%></TEXTAREA>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Actual ManHour</td>
      <td>
		<input name="ActManHour" value="<%=objReqDet.ActManHour%>" maxlength=3 style="WIDTH: 116px; HEIGHT: 20px" size=16
		maxlength="4" onKeyPress="if(!((window.event.keyCode >= '48')&&(window.event.keyCode <= '57'))){alert('Please enter NUMERIC values only.');return false;}"
		onblur="javascript:GetCostAct('<%=objReqDet.ActHourRate%>');">  hrs
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Actual Cost/Hour</td>
      			<td> <INPUT type=hidden name=ActHourRate value="<%=objReqDet.ActHourRate %> ">
      			<%=objReqDet.ActHourRate %></td>
    </tr>
    <tr>
      <td width="148" class="lbl1">  Actual Total Cost  </td>
      <td><INPUT name=ActTotalCost value="<%=objReqDet.ActTotalCost%>"
            maxlength=8 style="WIDTH: 116px; HEIGHT: 20px" size=16 >
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Act.&nbsp;Completed Date</td>
      <td><INPUT name=ActEndDate value="<%=objReqDet.ActEndDatestr%>"
            style="WIDTH: 115px; HEIGHT: 20px" size=16 >&nbsp;
			<a href="javascript:show_calendar('frmEdit.ActEndDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
			<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
      </td>
    </tr>
    </table>

    </td></tr>
</table></TD></TR></TABLE><!-- r3 end-->

<BR><!-- r4 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
	<table width="550" border="0" cellspacing="0" cellpadding="5">
      <tr><td class="trHdr"> QC's Update</td></tr>
	</table>
	</td></tr>
	<tr><td>

	<table width="550" border="0" cellspacing="2" cellpadding="5">
    <tr>
      <td width="148" class="lbl1" > Actual QC</td>
      <td><INPUT name=ActQC value="<%=objReqDet.ActQC%>" id=ActQC readonly style="WIDTH: 262px; HEIGHT: 20px" size=36 >
          <input name="cmdActQC" type="button" id="cmdActQC" value="..." onClick="javascript:showPopUp('ActQC');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">QC Status</td>
      <td><SELECT name="StatusQC" id="StatusQC" style="WIDTH: 65px" >
			  <OPTION value=""></OPTION>
              <OPTION value="P">Pass</OPTION>
              <OPTION value="F">Fail</OPTION></SELECT>
              <script>                  SetItemValue("StatusQC", "<%=objReqDet.StatusQC%>");</script>

      </td>


    </tr>
    <tr>
		<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
		<td><TEXTAREA name=RemarksQC id=RemarksQC style="WIDTH: 323px; HEIGHT: 77px"  rows=5 cols=50><%=objReqDet.RemarksQC%></TEXTAREA> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1" >UAT Ready Date</td>
      <td><INPUT name="UATReadyDate" value="<%=objReqDet.UATReadyDatestr%>" readonly>&nbsp;
			<a href="javascript:show_calendar('frmEdit.UATReadyDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
			<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1" >No of UAT Failed Status</td>
      <td><INPUT name="NoUATFailed" value="<%=objReqDet.NoUATFailed%>" style="WIDTH: 50px; HEIGHT: 20px"Readonly>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<b>UAT Failed</b>&nbsp;<INPUT name="UATFailed" value="<%=objReqDet.IsUATFailed%>" style="WIDTH: 50px; HEIGHT: 20px"Readonly>
      </td>
    </tr>
    </table>
    </td></tr>
</table></TD></TR></TABLE><!-- r4 end-->

<BR><!-- r5 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
	  <table width="550" border="0" cellspacing="0" cellpadding="5">
      <tr><td class="trHdr"> UAT Details</td></tr>
	</table>
	</td></tr>
	<tr><td>

	<table width="550" border="0" cellspacing="2" cellpadding="5"><TBODY>
    <tr>
      <td width="148" class="lbl1" > UAT User</td>
      <td><INPUT name=UserUAT value="<%=objReqDet.UserUAT%>" id=UserUAT readonly style="WIDTH: 262px; HEIGHT: 20px" size=36 >
          <input name="cmdUserUAT" type="button" id="cmdUserUAT" value="..." onClick="javascript:showPopUp('UserUAT');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">UAT Status</td>
      <td><SELECT name="StatusUAT" style="WIDTH: 65px" >
              <OPTION ></OPTION>
              <OPTION value="P">Pass</OPTION>
              <OPTION value="F">Fail</OPTION></SELECT>
              <script>                  SetItemValue("StatusUAT", "<%=objReqDet.StatusUAT%>");</script>
      </td>
    </tr>
    <tr>
		<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
		<td><TEXTAREA name=RemarksUAT id=RemarksUAT style="WIDTH: 323px; HEIGHT: 77px" rows=5 cols=50><%=objReqDet.RemarksUAT%></TEXTAREA> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1"> Expected Cut-in Date</td>
      <td><INPUT id="ExpCutinDate" name="ExpCutinDate" value="<%=objReqDet.ExpCutinDateStr%>" style="WIDTH: 115px; HEIGHT: 20px" size=16 >&nbsp;
			<a href="javascript:show_calendar('frmEdit.ExpCutinDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
			<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
	  </td>
    </tr>
    </td></tr>
</table></TD></TR></TBODY></TABLE><!-- r5 end-->

<BR><!-- r6 start-->
<table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
    <table width="550" border="0" cellspacing="0" cellpadding="5">
         <tr><td class="trHdr"> Deploy </td></tr>
	</table>
	</td></tr>
	<tr>
	<td>

	<table width="550" border="0" cellspacing="2" cellpadding="5">
    <tr>
      <td width="148" class="lbl1" > Deployed&nbsp;By</td>
      <td><INPUT name=DeployUser value="<%=objReqDet.DeployUser%>" id="DeployUser" readonly style="WIDTH: 262px; HEIGHT: 20px" size=36 >
          <input name="cmdDeployUser" type="button" id="cmdDeployUser" value="..." onClick="javascript:showPopUp('DeployUser');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Actual&nbsp;Cut-in Date</td>
      <td><INPUT name=ActCutinDate value="<%=objReqDet.ActCutinDateStr%>" id="ActCutinDate"
            style="WIDTH: 115px; HEIGHT: 20px" size=16 >
          <a href="javascript:show_calendar('frmEdit.ActCutinDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
			<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
      </td>
    </tr>
    <tr>
		<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
		<td><TEXTAREA name=RemarksDeploy id=RemarksDeploy style="WIDTH: 323px; HEIGHT: 77px"  rows=5 cols=50><%=objReqDet.RemarksDeploy%></TEXTAREA> </td>
    </tr>
    </table>
    </td></tr>
</table></TD></TR></TABLE><!-- r6 end-->

<BR><!-- r7 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr>
		<td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr"> Attachments </td>
			<td align="right" class="trHdr" style="WIDTH: 80px"></td>
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
		<br>
		<table width="550" border="1" cellspacing="1" cellpadding="2">
		<tr>
			<td class="lbl1">DEL</td>
			<td class="lbl1">File Name</td>
			<td class="lbl1">Added By</td>
			<td class="lbl1">Added Date</td>
		</tr>

<%
	introw=1
	If pageAction  = "" Then
		Dim rsfiles
		set rsfiles = objReqDet.RetrieveAttachements()

		'SHOW RS CONTENTS'
		'DocRefNo, ReqNo, DocFileName, CrUser, CrDate, ModUser, ModDate '
		'  0         1       2          3        4       5        6     '
		Do While rsfiles.EOF = False
		'if rsfiles.EOF = False then
%>
		<tr>
		<%if Trim(objReqDet.SCode)="Cancelled" or Trim(objReqDet.SCode)="Closed" then %>
			<td width="30" class="">&nbsp; </td>
		<% Else %>

			<td width="30" class=""><a href="javascript:DeleteRow(<%=introw%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>

			</td>
		<%End If %>
			<td style="WIDTH: 120px" width=120><a href="../upload/<%=rsfiles(0)%>"><%=rsfiles(2)%></a>
				<input type="hidden" name="fname" value="<%=rsfiles(0)%>">
				<input type="hidden" name="fnamedis" value="<%= rsfiles(2)%>">
				</td>
			<td style="WIDTH: 120px" width=120><%=rsfiles(3)%>
				<input type="hidden" name="fuser" value="<%=rsfiles(3)%>"></td>
			<td style="WIDTH: 120px" width=120><%=rsfiles(4)%>
				<input type="hidden" name="fdate" value="<%=rsfiles(4)%>"></td>
				
		</tr>
<%

		introw=introw+1

		rsfiles.MoveNext
		Loop%>
		<script>		
		<% Select case Trim(objReqDet.SCode)  %>

		<% CASE "Open" %>
			frmEdit.RemarksIT.focus();

		<% CASE "Scoping" %>
			frmEdit.RemarksIT.focus();

		<% CASE "Queued" %>
			frmEdit.StatusQC.focus();

		<% CASE "DIP" %>
			if (frmEdit.ActEndDate.value=="")
				frmEdit.RemarksDev.focus();
			else
				frmEdit.RemarksUAT.focus();
		<% CASE "UAT" %>
			frmEdit.ExpCutinDate.focus();

		<% CASE "Deploy" %>
			frmEdit.RemarksDeploy.focus();

		<% END Select %>
		</script>

	<%Else
	introw=0
		if pageAction = "DEL" THEN
		
			arrcount = Request.Form("fname").Count
			for i = 2 to arrcount
				If cint(i-1) <> cint(delIndex) then
		%>
				<tr>

					<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=introw+1%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>

					<td style="WIDTH: 120px" width=120><a href="../upload/<%=Request.Form("fname")(i)%>"><%=Request.Form("fnamedis")(i)%></a>
						<input type="hidden" name="fname" value="<%=Request.Form("fname")(i)%>">
						<input type="hidden" name="fnamedis" value="<%=Request.Form("fnamedis")(i)%>"></td>
					<td style="WIDTH: 120px" width=120><%=Request.Form("fuser")(i)%>
						<input type="hidden" name="fuser" value="<%=Request.Form("fuser")(i)%>"></td>
					<td style="WIDTH: 120px" width=120><%=Request.Form("fdate")(i)%>
						<input type="hidden" name="fdate" value="<%=Request.Form("fdate")(i)%>"></td>
				</tr>
				<%introw=introw+1
				end if
			next
		ELSE
			arrcount = Request.Form("fname").Count
			for i = 2 to arrcount	%>
				<tr>
				<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=introw+1%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>
				<td style="WIDTH: 120px" width=120><a href="../upload/<%=Request.Form("fname")(i)%>"><%=Request.Form("fnamedis")(i)%></a>
					<input type="hidden" name="fname" value="<%=Request.Form("fname")(i)%>">
					<input type="hidden" name="fnamedis" value="<%=Request.Form("fnamedis")(i)%>"></td>
				<td style="WIDTH: 120px" width=120><%=Request.Form("fuser")(i)%>
					<input type="hidden" name="fuser" value="<%=Request.Form("fuser")(i)%>"></td>
				<td style="WIDTH: 120px" width=120><%=Request.Form("fdate")(i)%>
					<input type="hidden" name="fdate" value="<%=Request.Form("fdate")(i)%>"></td>
				</tr>
				<%introw=introw+1
			next
		END IF
	End If

	if (pageAction = "UPLOAD") then
%>
		<tr>
			<td width="30" class="trDet"><a href="javascript:DeleteRow(<%=introw+1%>);" ><img src="../images/delete.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>	</td>
			<td style="WIDTH: 120px" width=120><a href="../upload/<%=Request.Form("fname1")%>"><%=Request.Form("fname1dis")%></a>
				<input type="hidden" name="fname" value="<%=Request.Form("fname1")%>">
				<input type="hidden" name="fnamedis" value="<%=Request.Form("fname1dis")%>"></td>
			<td style="WIDTH: 120px" width=120><%=strUser%>
				<input type="hidden" name="fuser" value="<%=strUser%>"></td>
			<td style="WIDTH: 120px" width=120><%=now()%>
				<input type="hidden" name="fdate" value="<%=Now()%>"></td>
		</tr>
<%
	end if
%>
		</table>


		</td>
	</tr>
</table>
	<%if Trim(objReqDet.SCode)="Cancelled" or Trim(objReqDet.SCode)="Closed" then %>

	<%elseif Trim(objReqDet.SCode)="Hold" then %>

	<% Else  %>
		<br>
		<table width="550" border="1" cellspacing="0" cellpadding="5" borderColor=salmon>
			<tr>

			<td align="right"> <INPUT name=cmdAttach id=cmdAttach type=button size=7 value="Attachments"  style="WIDTH: 110px; HEIGHT: 22px" onClick="javascript:attach();">
			</td>
			</tr>
		</table>
	<% End if %>
</TD></TR></TABLE><!-- r7 end-->
</form>
</body>

</html>
