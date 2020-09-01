<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dim datefr, dateto, datefrmgr, datetomgr
dim intreqno
%>
<html>
<head>
<Title>ASR - Search</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->

<script language="javascript" src="datepicker.js"></script>
<script>
var popWin;
function showPopUp(skey) {
	//if (popwin != null) { popWin.close() }
	var url = "poplist.asp" + "?skey=" + skey;
	popWin = open(url, "popupwin", "toolbar=no,scrollbars=yes,width=480,height=300,top=250, left=200");
}

function returnCustomer(customer) {
	frmSearch.customer.value = customer;
}

function returnAppl(appl) {
	frmSearch.appl.value = appl;
}
function returnAppMgr(ipvalue)
{
	frmSearch.appmgr.value = ipvalue;
}

function returnTeamLead(TeamLead)
{
	frmSearch.teamlead.value = TeamLead;
}
function returnDeveloperList(ipvalue)
{
	frmSearch.developer.value = ipvalue;
}
function returnReqUserList(userid)
{
	frmSearch.requestor.value = userid;
}
function showDets(reqno, status)
{
if ((status=="Closed") || (status=="Cancelled"))
{
	var url = "viewrequest.asp?ReqNo=" + reqno;
	popWin = open(url, "ViewReq", "toolbar=yes,resizable=yes,scrollbars=yes,width=600,height=450,top=60, left=100");
}
else
{
	var url = "editrequest.asp?ReqNo=" + reqno;
	//var url = "sd.asp?ReqNo=" + reqno;
	popWin = open(url, "EditReq", "toolbar=yes,resizable=yes,scrollbars=yes,width=600,height=450,top=60, left=100");
	//popWin = open(url, "EditReq", "toolbar=yes,resizable=yes,scrollbars=yes,width=screen.availWidth,height=screen.availHeight,top=60, left=100");
	//alert(screen.availHeight);
}
}
function CostForm(reqno)
{
	var url = "requestcost.asp?reqno="+ reqno;
	popWin = open(url, "MonthlyCostForm", "toolbar=yes,resizable=yes,scrollbars=yes,width=600,height=550,top=60, left=100");
}
</script>

<%
dim mobj, rs, strList
dim pageAction, reqno
dim buf, skey, i
dim rowCount

' Check to see if there is value in the NAV querystring.  If there
	' is, we know that the client is using the Next and/or Prev hyperlinks
	' to navigate the recordset.
	If Request.QueryString("NAV") = "" Then
		intPage = 1
	Else
		intPage = Request.QueryString("NAV")
	End If

pageAction = Request("pageAction")
reqno = Request("reqno")
Set mobj = Server.CreateObject("ASRTrans.clsList")
dateto = mobj.RetrieveCostEntryDtTo()
datefr = mobj.RetrieveCostEntryDtFrom()
datetomgr = mobj.RetrieveCostEntryDtToMgr()
datefrmgr = mobj.RetrieveCostEntryDtFromMgr()
set rs=mobj.RetrieveStatus()
'NEW function to return <option list as STRING'
strList = "<option value=""""> </option><br>"
do while rs.eof=false
	strList =  strList & "<option value=""" & rs.fields(0) & """>" & rs.fields(0) &  "</option><br>"
	rs.movenext
loop

%>
<form method="post" name=frmSearch>
<input type="hidden" name="pageAction" value="SEARCH">
  <table width="500" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966CC">
   <tr><td>

	<table width="500" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td class="trHdr">Search Documents</td>
			<td align="right" class="trHdr">
			  <input name="cmdSubmit" type="submit" id="cmd2" value="Search">

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
        <input type="text" name="appl" value="<%=Request("appl")%>"> <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPL');">
      </td>
    </tr>
    <tr>
      <td width="148" class="lbl1">Request No</td>
      <td>
        <input type="text" name="reqno" value=<%=Request("reqno")%>>
      </td>
    </tr>
   <tr>
      <td class="lbl1">Status</td>
      <td> <select name="status"> <%=strList %> </select> </td>
    </tr>
 <script>
document.forms[0].status.value = "<%= Request("status") %>";
</script>
    <tr>
      <td width="148" class="lbl1">IT Account Manager</td>
      <td>
        <input type="text" name="appmgr" value="<%=Request("appmgr")%>"> <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('APPMGR');">
      </td>
    </tr>
    <tr>
	      <td width="148" class="lbl1">Team Leader</td>
	      <td>
	        <input type="text" name="teamlead" value="<%=Request("teamlead")%>">
	        <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('TeamLead');">
	      </td>
    </tr>
    <tr>
	      <td width="148" class="lbl1">Developer</td>
	      <td>
	        <input type="text" name="developer" value="<%=Request("developer")%>"> <input name="cmd1" type="button" id="cmd1" value="..." onClick="javascript:showPopUp('DeveloperList');">
	      </td>
    </tr>
    <tr>
	      <td width="148" class="lbl1">Requestor</td>
	      <td>
	        <input type="text" name="requestor" value="<%=Request("requestor")%>">
	        <input name="cmdRequestor" type="button" id="cmdRequestor" value="..." onClick="javascript:showPopUp('USERLIST');">
	      </td>
    </tr>
    <tr>
		<td class="lbl1">&nbsp;Request Date From: <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		To:</td>
		<td><input type="text" name="fromdate" value="<%=Request("fromdate")%>" size="12" maxlength="15">
		<a href="javascript:show_calendar('frmSearch.fromdate');" onMouseOver="window.status='Select Date';return true;" onMouseOut="window.status='';return true;">
		<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
		<br><input type="text" name="todate" value="<%=Request("todate")%>" size="12" maxlength="15">
		<a href="javascript:show_calendar('frmSearch.todate');" onMouseOver="window.status='Select Date';return true;" onMouseOut="window.status='';return true;">
		<img src="../images/calendar.gif" border=0 width="22" height="19" border="0" align="absmiddle"></a>
		</td>
	</tr>
	
	<tr>
	    <td width="148" class="lbl1">Request Type</td>
	    <td><select name="reqtype">
			<option value="" selected></option>
			<option value="R"> Request</option>
			<option value="P">Project</option>
        </select></td>

<script>
<% if Request("reqtype") <> "" then %>document.forms[0].reqtype.value = "<%= Request("ReqType") %>"; <% end if %>
</script>


    </tr>
    <tr>
      <td class="lbl1">Sort By</td>
      <td><select name="sortby">
			<option value="ApplName"> Application</option>
			<option value="CName">Customer</option>
			<option value="DeptName">Dept </option>
			<option value="Developer">Developer</option>
			<option value="ExpEndDate"> ExpCompleteDate</option>
			<option value="ExpStartDate">ExpStartDate </option>
			<option value="AppMgr">IT Account Manager</option>
			<option value="Priority"> Priority</option>
			<option value="ReqNo" selected>RequestNo </option>
			<option value="ReqUser">Requestor</option>
			<option value="SCode">Status</option>
			<option value="TeamLead">Team Leader</option>
        </select></td>
 <script>
<% if Request("sortby") <> "" then %>document.forms[0].sortby.value = "<%= Request("sortby") %>"; <% end if %>
</script>

      <td><select name="order">
			<option value="ASC"> Ascending</option>
			<option value="DESC" selected>Descending</option>
        </select></td>

<script>
<% if Request("order") <> "" then %>document.forms[0].order.value = "<%= Request("order") %>"; <% end if %>
</script>


    </tr>
  </table>

</td></tr>
</table>
</form>

	<%

	if (pageAction = "SEARCH") then

		Set mobj = Server.CreateObject("ASRTrans.clsSearch")

		if Request("reqno")="" then
			intreqno=0
		else
			intreqno=Clng(Request("reqno"))
		end if

		set rs=mobj.RetrieveSearch(Request("customer"),Request("appl"),Request("status"),Request("sortby"), _
					Request("order"), intreqno, Request("developer"), Request("appmgr"), Request("teamlead"), _
					Request("requestor"), Request("fromdate"), Request("todate"), Request("reqtype"))
		rs.PageSize = 50
		rs.CacheSize = rs.PageSize
		intPageCount = rs.PageCount
		intRecordCount = rs.RecordCount
'response.write 	intRecordCount
		If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
		If CInt(intPage) <= 0 Then intPage = 1

		' Make sure that the recordset is not empty.  If it is not, then set the
		' AbsolutePage property and populate the intStart and the intFinish variables.
		If intRecordCount > 0 Then
			rs.AbsolutePage = intPage
			intStart = rs.AbsolutePosition
			If CInt(intPage) = CInt(intPageCount) Then
				intFinish = intRecordCount
			Else
				intFinish = intStart + (rs.PageSize - 1)
			End if
		End If

	%>

	<% if rs.eof=false then %>
	<table>
			<tr >
		<%
			' Check to see if the current page is greater than the first page
			' in the recordset.  If it is, then add a "Previous" link.

			If cInt(intPage) > 1 Then	%>
		   		<td><a href="search.asp?pageAction=SEARCH&NAV=<%=intPage - 1%>"><< Prev</a></td>
		   		<td width=10></td>
			<%End IF%>
			<%
			' Check to see if the current page is less than the last page
			' in the recordset.  If it is, then add a "Next" link.
			If cInt(intPage) < cInt(intPageCount) Then	%>
		   		<td colspan=2  align="center" ><a href="search.asp?pageAction=SEARCH&NAV=<%=intPage + 1%>">Next >></a></td>
			<%End If%>
		</tr>
		</table>
		<br>
		<br>

  		<table width="800" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
			<tr>
			
				<%If (appUserType="I" and day(date)>=datefr and day(date)<=dateto) or _
					  (appAccessLevel="1" and day(date)>=datefrmgr and day(date)<=datetomgr) Then  %>
						<th class="trHdr" width=50 >&nbsp;Cost&nbsp;</th>								<%End If %>	
				<% for i=0 to rs.fields.Count-2 %>
				<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
				<% next %>
			</tr>

			<%
			' Iterate through the recordset until we reach the end of the page
			' or the last record in the recordset.
			For intRecord = 1 to rs.PageSize %>
				<tr>
					
					<%If (appUserType="I" and day(date)>=datefr and day(date)<=dateto) or  _
						(appAccessLevel="1" and day(date)>=datefrmgr and day(date)<=datetomgr) Then  %>
					<td  class="" align="center" valign="center">
					<%If trim(rs.fields("Status")) <>"Closed" and _
						 trim(rs.fields("Status")) <>"Open" and trim(rs.fields("ReqType")) = "P" Then %>
					<A href="javascript:CostForm('<%=rs.fields("ReqNo")%>');"  ><img src="../images/money.gif" border=0 width="25" height="25" border="0" align="absmiddle"></A>	
					<%Else%>	
					&nbsp;
					<%End If%>					
					</td>
					<%End If%>
					<td class="" valign="top">&nbsp;<a href="javascript:showDets('<%=rs.fields(0)%>', '<%=trim(rs.fields(7))%>');"><%=rs.fields(0).value%></a>&nbsp;
					<br>
					&nbsp;&nbsp;&nbsp;<%= rs.fields("ReqType").value %>
					</td>
					<%for i=1 to rs.fields.Count-2%>
						<td class="" valign="top">&nbsp;<%= rs.fields(i).value %>&nbsp;</td>
					<%next %>
				</tr>
				<% rs.MoveNext
				If rs.EOF Then Exit for
			Next %>

			</table>
			<br>
			<br>
		<table>
		<tr >
	<%
		' Check to see if the current page is greater than the first page
		' in the recordset.  If it is, then add a "Previous" link.

		If cInt(intPage) > 1 Then	%>
	   		<td><a href="search.asp?pageAction=SEARCH&NAV=<%=intPage - 1%>"><< Prev</a></td>
	   		<td width=10></td>
		<%End IF%>
		<%
		' Check to see if the current page is less than the last page
		' in the recordset.  If it is, then add a "Next" link.
		If cInt(intPage) < cInt(intPageCount) Then	%>
	   		<td colspan=2  align="center" ><a href="search.asp?pageAction=SEARCH&NAV=<%=intPage + 1%>">Next >></a></td>
		<%End If%>
	</tr>
		</table>
		<%

		rs.close
	else%>
		There is no data to show
	<%End if
	Set rs = nothing
	Set mobj = nothing
end if 'search action'
%>



	<br>
	<br>
<a href="javascript:window.scrollTo(0,0);">Go to Top</a>

</form>
<!-- #include file="footer.asp" -->
</body>
</html>
