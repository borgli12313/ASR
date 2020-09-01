<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<Title>ASR - Search</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>

<body>
<!-- #include file="links.asp" -->
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
	popWin = open(url, "EditReq", "toolbar=yes,resizable=yes,scrollbars=yes,width=600,height=450,top=60, left=100");
}
}
</script>

<%
dim mobj, rs, strList
dim pageAction, reqno
dim buf, skey, i
dim rowCount

pageAction = Request("pageAction")
reqno = Request("reqno")
Set mobj = Server.CreateObject("ASRTrans.clsList")
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
			<option value="DESC">Descending</option>
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
			set rs=mobj.RetrieveSearch(Request("customer"),Request("appl"),Request("status"),Request("sortby"),Request("order"), 0, Request("developer"), Request("appmgr"), Request("teamlead"), Request("requestor"))
		Else
			set rs=mobj.RetrieveSearch(Request("customer"),Request("appl"),Request("status"),Request("sortby"),Request("order"), Clng(Request("reqno")), Request("developer"), Request("appmgr"), Request("teamlead"), Request("requestor"))
		End if
	%>

	<% if rs.eof=false then

		If rs.recordcount>100 Then %>
			The search result exceeded the Maximum number of rows. Please apply some search condition and try again!
		<%Else%>
  		<table width="800" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
			<tr>
				<% for i=0 to rs.fields.Count-1 %>
				<th class="trHdr">&nbsp;<%= rs.fields(i).name %>&nbsp;</th>
				<% next %>
			</tr>
			<%do while rs.eof=false%>
			<tr>
				<td class="" valign="top">&nbsp;<a href="javascript:showDets('<%=rs.fields(0)%>', '<%=trim(rs.fields(7))%>');"><%=rs.fields(0).value%></a>&nbsp;
				</td>
				<%for i=1 to rs.fields.Count-1%>
					<td class="" valign="top">&nbsp;<%= rs.fields(i).value %>&nbsp;</td>
				<%next %>

			</tr>
			<%
			rs.movenext
			loop
		End if
		rs.close
	else%>
		There is no data to show
	<%End if
	Set rs = nothing
	Set mobj = nothing
end if 'search action'
%>


	</table>
	<br>
	<br>
<a href="javascript:window.scrollTo(0,0);">Go to Top</a>


</body>
</html>
