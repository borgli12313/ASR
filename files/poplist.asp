<html>
<head>
<title>Popup Listing</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>
<script>

function returnAppl(appl)
{
	opener.returnAppl(appl);
	window.close();
}
function returnCustomer(customer)
{
	opener.returnCustomer(customer);
	window.close();
}

function returnCustAppl(customer,appl,appmgr)
{
	opener.returnCustAppl(customer,appl,appmgr);
	window.close();
}
function returnReqUser(userid, email, dept, mgr, mgremail)
{
	opener.returnReqUser(userid, email, dept, mgr, mgremail);
	window.close();
}

function returnReqUserList(userid)
{
	opener.returnReqUserList(userid);
	window.close();
}

function returnReqUserMgr(mgr, mgremail)
{
	opener.returnReqUserMgr(mgr, mgremail);
	window.close();
}

function returnAppMgr(ipvalue)
{
	opener.returnAppMgr(ipvalue);
	window.close();
}
function returnAppBkpMgr(ipvalue)
{
	opener.returnAppBkpMgr(ipvalue);
	window.close();
}
function returnTeamLead(TeamLead)
{
	opener.returnTeamLead(TeamLead);
	window.close();
}
function returnDeveloperList(ipvalue)
{
	opener.returnDeveloperList(ipvalue);
	window.close();
}
function returnDeveloper()
{
	var chksel = false;
	if (! document.frmTest.chk) return;

	chksel = isSelected(document.frmTest.chk); 
	if (!chksel) {
		alert("You must select atleast one value!");
		return;
	}

	var arrlen = document.getElementsByTagName("INPUT").length; 
	var strDev="";
	if (document.getElementsByTagName("INPUT")[1].checked)
	{
	strDev =document.getElementsByTagName("INPUT")[1].value ;
	}
	else
	{
		for (i = 2; i < arrlen; i++)
		{
		if (document.getElementsByTagName("INPUT")[i].checked)
		strDev =(strDev + document.getElementsByTagName("INPUT")[i].value + ", ") ;
		}
		strDev=strDev.substring(0,strDev.length-2);
	}
	//alert(strDev);


	opener.returnDeveloper(strDev);
	window.close();
}
function isSelected(obj) {
	var chkflg = false;

	if(! obj) return chkflg;
	var chkcount = obj.length;

	if(chkcount == 1) chkflg = obj.checked;

	if(chkcount > 1) {
		for(var i = 0; i < chkcount; i++){
			chkflg = obj[i].checked ;
			if(chkflg) break;
		}
	}
	return chkflg;
}

function returnUserUAT(UserUAT)
{
	opener.returnUserUAT(UserUAT);
	window.close();
}
function returnActQC(ActQC)
{
	opener.returnActQC(ActQC);
	window.close();
}
function returnExpQC(ExpQC)
{
	opener.returnExpQC(ExpQC);
	window.close();
}
function returnDeployUser(deployuser)
{
	opener.returnDeployUser(deployuser);
	window.close();
}

</script>

<%
dim mobj, rs, buf, skey

skey = Request("skey")

Set mobj = Server.CreateObject("ASRTrans.clsList")%>
<form method="post"  id=form1 name=frmTest action="">
<%select case skey

	case "APPL"
		set rs=mobj.RetrieveAppl()
			%>
			<table border="1">
			<tr>
			<th><%= rs.fields(0).name %></th>
			</tr>
			<%
			do while rs.eof=false
			%>
				<tr>
				<td><a href="javascript:returnAppl('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
				</tr>
			<%
				rs.movenext
			loop

	case "CUSTOMER"
		set rs=mobj.RetrieveCustomer()
		%>
		<table border="1">
		<tr>
		<th><%= rs.fields(0).name %></th>
		</tr>
		<%
			do while rs.eof=false
		%>
		<tr>
		<td><a href="javascript:returnCustomer('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
		</tr>
		<%
				rs.movenext
			loop

	case "CUSTAPPL"
		set rs=mobj.RetrieveCustApplMgr()
		'set rs=mobj.RetrieveUser()
		%>

		<table border="1">
		<tr>
		<th><%= rs.fields(0).name %></th>
		<th><%= rs.fields(1).name %></th>
		<th><%= rs.fields(2).name %></th>
		</tr>
		<%	do while rs.eof=false	%>
			<tr>
			<td><a href="javascript:returnCustAppl('<%= rs.fields(0) %>','<%= rs.fields(1) %>','<%= rs.fields(2) %>');"><%= rs.fields(0) %></a></td>
			<td>&nbsp;<%= rs.fields(1) %></td>
			<td>&nbsp;<%= rs.fields(2) %></td>
			</tr>
			<%	rs.movenext
			loop


	case "USER"
		set rs=mobj.RetrieveUser()
		%>
		<table border="1">
		<tr>
		<th><%= rs.fields(0).name %></th>
		<th><%= rs.fields(1).name %></th>
		<th><%= rs.fields(2).name %></th>
		<th><%= rs.fields(3).name %></th>
		<th><%= rs.fields(4).name %></th>
		</tr>
		<%
		do while rs.eof=false
			%>
			<tr>
			<td><a href="javascript:returnReqUser('<%= rs.fields(0) %>','<%= rs.fields(1) %>','<%= rs.fields(2) %>','<%= rs.fields(3) %>','<%= rs.fields(4) %>');"><%= rs.fields(0) %></a></td>

			<td>&nbsp;<%= rs.fields(1) %></td>
			<td>&nbsp;<%= rs.fields(2) %></td>
			<td>&nbsp;<%= rs.fields(3) %></td>
			<td>&nbsp;<%= rs.fields(4) %></td>

			</tr>
			<%
			rs.movenext
		loop

	case "USERLIST"
		set rs=mobj.RetrieveUser()
		%>
		<table border="1">
		<tr>
		<th><%= rs.fields(0).name %></th>
		</tr>
		<%
		do while rs.eof=false
			%>
			<tr>
			<td><a href="javascript:returnReqUserList('<%= rs.fields(0)%>');"><%= rs.fields(0) %></a></td>

			</tr>
			<%
			rs.movenext
		loop

	case "USERMGR"
		set rs=mobj.RetrieveUserMgr()
		%>
		<table border="1">
		<tr>
		<th><%= rs.fields(0).name %></th>
		<th><%= rs.fields(1).name %></th>
		</tr>
		<%
		do while rs.eof=false
			%>
			<tr>
			<td><a href="javascript:returnReqUserMgr('<%= rs.fields(0) %>','<%= rs.fields(1) %>');"><%= rs.fields(0) %></a></td>

			<td>&nbsp;<%= rs.fields(1) %></td>
			</tr>
			<%
			rs.movenext
		loop

	case "APPMGR"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnAppMgr('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnAppMgr('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop
	case "APPBKPMGR"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnAppBkpMgr('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnAppBkpMgr('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop
	case "TeamLead"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnTeamLead('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnTeamLead('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop

case "DeveloperList"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnTeamLead('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnDeveloperList('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop
	case "Developer"
		set rs=mobj.RetrieveITUser()
		%>

		<table border="1" width=150>

		<tr width=100><th ><%= rs.fields(0).name %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT id=cmdOK type=button value=OK name=cmdOK style="WIDTH: 42px; HEIGHT: 21px" size=28 onClick="javascript:returnDeveloper();">
		</th></tr>
			<tr>
			<td><INPUT id=checkbox1 type=checkbox name=chk value=""> Clear Selection</td>
			</tr>
		<%	do while rs.eof=false %>
			<tr>
			<td><INPUT id=checkbox1 type=checkbox name=chk value="<%=rs.fields(0)%>"> <%= rs.fields(0) %></td>
			</tr>
			<%
			rs.movenext
		loop

	case "ExpQC"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnExpQC('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnExpQC('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop

	case "ActQC"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnActQC('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnActQC('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop

	case "UserUAT"
		set rs=mobj.RetrieveUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnUserUAT('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnUserUAT('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop

	case "DeployUser"
		set rs=mobj.RetrieveITUser()
		%>
		<table border="1">
		<tr><th><%= rs.fields(0).name %></th></tr>
			<tr>
			<td><a href="javascript:returnDeployUser('');">Clear Selection</a></td>
			</tr>

		<%	do while rs.eof=false %>
			<tr>
			<td><a href="javascript:returnDeployUser('<%= rs.fields(0) %>');"><%= rs.fields(0) %></a></td>
			</tr>
			<%
			rs.movenext
		loop

end select
Set rs = nothing
Set mobj = nothing
%>

</table>
</form>
</html>
