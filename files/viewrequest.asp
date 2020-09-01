<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Option Explicit
	Dim objReqDet, strUser, retVal, mobj, objReqCost
	Dim rs, deptList, msg, hrRate, dt, introw, rsReqCost
	Dim i, arrcount, pageAction, delFile, delIndex

	strUser=Request.ServerVariables ("LOGON_USER")
	'strUser = "hema"
	strUser = Mid(strUser,Instr(1,strUser,"\")+1)
	
	Set objReqDet = Server.CreateObject("ASRTrans.clsRequestDets")
	Set objReqCost = Server.CreateObject("ASRTrans.clsReqCost")
	objReqCost.ReqNo = Request("ReqNo")
	
	If strUser="" then
		msg = "Access Denied"
	Else
		pageAction = Request.Form("pageAction")

		If pageAction  = "" then
			objReqDet.ReqNo = Request("ReqNo")
			'Response.Write objReqDet.ReqNo
			On Error Resume Next
			objReqDet.RetrieveRecord
			If Err Then
				msg = "There is some error in getting your details please try again! <br>Error " & Err.Number & " : " & Err.Description
			End If
		End If '--pageAction =--'
	End If '-- Access Failed --'

	Set mobj = Server.CreateObject("ASRTrans.clsList")
	deptList = mobj.PopulateDept()
	dt = mobj.RetrieveDate()
	hrRate = mobj.RetrieveHourRate()

%>

<html>
<head>
<Title>ASR - View Request Details</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">


</head>

<!-- #include file="access2.asp" -->

<script language="Javascript" src="fnlist.js"></script>

<% if msg <> "" then %><div class="msginfo"><%=msg %></div><% end if %>
<form method="post"  id=form1 name=frmEdit action="editrequest.asp">
<INPUT type="hidden" name="pageAction" value="<%=pageAction%>">
<INPUT type="hidden" name="ReqNo" value="<%=Request("ReqNo")%>">
<INPUT type="hidden" name="SCode" value="<%=trim(objReqDet.SCode)%>">
<INPUT type="hidden" name="UID" value="<%=strUser%>">

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
				<td class="txtro"><%=objReqDet.ReqUser%> </td>
		</tr>
    <tr>
      <td width="148" class="lbl1">Email</td>
      <td class="txtro"><%=objReqDet.ReqEmail%></td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Requestor's Manager</td>
      <td class="txtro"><%=objReqDet.ReqUserMgr%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Email</td>
      <td class="txtro"><%=objReqDet.ReqMgrEmail%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Dept</td>
      <td class="txtro"> <%=objReqDet.deptname%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Customer</td>
      <td class="txtro"><%=objReqDet.CName%> </td>
    </tr>
	<tr>
      <td width="148" class="lbl1">Application</td>
      <td class="txtro"> <%=objReqDet.ApplName%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">IT Account Manager</td>
      <td class="txtro"> <%=objReqDet.AppMgr%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Priority</td>
      <td class="txtro"> <select name="priority"
            style="WIDTH: 140px" disabled>
          <option value="1">Critical</option>
          <option value="2">Important &amp; Urgent</option>
          <option value="3">Important</option>
          <option value="4">Nice to have</option>
          <option value="5">Suggestion</option>
        </select>
        <script>SetItemValue("priority","<%=objReqDet.Priority%>");</script>
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Request Title</td>
      <td class="txtro">  <%=objReqDet.ReqTitle%></td>
    </tr>
    <tr>
      <td class="lbl1" style="VERTICAL-ALIGN: top">Description</td>
    <td class="txtro"> <%=objReqDet.ReqDesc%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Expected Go Live Date</td>
      <td class="txtro"> <%=objReqDet.ExpCloseDatestr%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Request Date</td>
      <td class="txtro"><%=objReqDet.CrDate%> </td>
    </tr>
    </table>

	</td></tr>
  </table><!-- r1 end-->

<BR><!-- rem start-->
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
			<td class="txtro">&nbsp;  <%=objReqDet.RemarksCH%> </td>
		</tr>
		</table>
    </td></tr>
  </table><!-- rem end-->

<BR><!-- r2 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="550" border="0" cellspacing="0" cellpadding="5">
			<tr><td class="trHdr"> IT Details</td>
		    </tr>
		</table>
	</td></tr>
	<tr><td>
		<table width="550" border="0" cellspacing="2" cellpadding="5">
			<tr>
				<td width="148" class="lbl1">Team Leader</td>
				<td class="txtro"> &nbsp;<%=objReqDet.TeamLead%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Team Members</td>
				<td class="txtro">  &nbsp;<%=objReqDet.Developer%>  </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Quality Control</td>
				<td class="txtro"> &nbsp;<%=objReqDet.ExpQC%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Estimated ManHour</td>
				<td class="txtro"> <%=objReqDet.EstManHour%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Estimated Cost/Hour</td>
					<td class="txtro"> <%=hrRate%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Estimated Total Cost</td>
				<td class="txtro"> <%=objReqDet.EstTotalCost%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Exp. Start Date</td>
				<td class="txtro"> &nbsp;<%=objReqDet.ExpStartDatestr%> </td>
			</tr>
			<tr>
				<td width="148" class="lbl1">Exp. Close Date</td>
				<td class="txtro"> &nbsp;<%=objReqDet.ExpEndDatestr%> </td>
			</tr>
			<tr>
				<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
				<td class="txtro"> &nbsp;<%=objReqDet.RemarksIT%>  </td>
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
		end if
		rsReqCost.close
		set rsReqCost = Nothing
		
%>
</table>
</td></tr>
</table>
<!-- cost end-->
<%End If%>
<BR><!-- r3 start-->
  <table width="550" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
  <table width="550" border="0" cellspacing="0" cellpadding="5">
          <tr>
      <td class="trHdr"> Programmer's Update</td>

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
      <td class="txtro"> &nbsp;<%=objReqDet.ActStartDateStr%> </td>
    </tr>
    </table>
    <HR>
    <table width="550" border="0" cellspacing="2" cellpadding="5">
    <tr>
      <td width="148" class="lbl1" style="VERTICAL-ALIGN: top">Progress Details</td>
      <td class="txtro">&nbsp;<%=objReqDet.RemarksDev%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Actual ManHour</td>
      <td class="txtro"> <%=objReqDet.ActManHour%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Actual Cost/Hour</td>
      			<td class="txtro"> <%=hrRate%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">  Actual Total Cost  </td>
      <td class="txtro"> <%=objReqDet.ActTotalCost%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Act.&nbsp;Completed Date</td>
      <td class="txtro"> &nbsp;<%=objReqDet.ActEndDatestr%> </td>
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
      <td class="txtro"> &nbsp;<%=objReqDet.ActQC%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">QC Status</td>
      <td class="txtro"><SELECT name="StatusQC" id="StatusQC" style="WIDTH: 65px" disabled>
			  <OPTION value=""></OPTION>
              <OPTION value="P">Pass</OPTION>
              <OPTION value="F">Fail</OPTION></SELECT>
              <script>SetItemValue("StatusQC","<%=objReqDet.StatusQC%>");</script>
      </td>


    </tr>
    <tr>
		<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
		<td class="txtro">&nbsp; <%=objReqDet.RemarksQC%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1" >UAT Ready Date</td>
      <td class="txtro"> &nbsp;<%=objReqDet.UATReadyDatestr%> </td>
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
      <td class="txtro"> &nbsp;<%=objReqDet.UserUAT%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">UAT Status</td>
      <td class="txtro"><SELECT name="StatusUAT" style="WIDTH: 65px" disabled>
              <OPTION ></OPTION>
              <OPTION value="P">Pass</OPTION>
              <OPTION value="F">Fail</OPTION></SELECT>
              <script>SetItemValue("StatusUAT","<%=objReqDet.StatusUAT%>");</script>
      </td>
    </tr>
    <tr>
		<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
		<td class="txtro">&nbsp; <%=objReqDet.RemarksUAT%>  </td>
    </tr>
    <tr>
      <td width="148" class="lbl1"> Expected Cut-in Date</td>
      <td class="txtro"> &nbsp;<%=objReqDet.ExpCutinDateStr%> </td>
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
      <td class="txtro"> &nbsp;<%=objReqDet.DeployUser%> </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Actual&nbsp;Cut-in Date</td>
      <td class="txtro"> &nbsp;<%=objReqDet.ActCutinDateStr%> </td>
    </tr>
    <tr>
		<td class="lbl1" style="VERTICAL-ALIGN: top">Remarks</td>
		<td class="txtro">&nbsp; <%=objReqDet.RemarksDeploy%> </td>
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
	introw=0
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
			<td width="30" class="">&nbsp; </td>
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

		introw=introw=+1
		rsfiles.MoveNext
		Loop

	End If%>


		</table>


		</td>
	</tr>
</table>
</TD></TR></TABLE><!-- r7 end-->

</form>
</body>
</html>
