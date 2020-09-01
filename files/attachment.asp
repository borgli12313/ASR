<html>
<head>
<title>ASR : File Attachments</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/app.css" rel="stylesheet" type="text/css">
</head>
<script>
function returnFile(fname,fnamedis)
{
	opener.returnFile(fname,fnamedis);
	window.close();
}

function submitform()
{
	var frmobj = document.forms[0];
	if(frmobj.fname1.value == "")
	{
	alert("please enter the file name.");
    frmobj.fname1.focus();
    return false;
	}
	var fn = frmobj.fname1.value;
	var fpos = fn.lastIndexOf("\\") + 1;
	fn=fn.substr(fpos);
	
	if(fn.length > 50)
	{
	alert("The file name exceeded the maximum limit (50 characters). please change the file name and try again.");
    frmobj.fname1.focus();
    return false;
	}
	
	frmobj.submit();
}
</script>

<%
if (Request("err") = "1") Then
%>
<div class="msginfo">There is some error in attaching the file. Please try again!
<BR><%=Request("msg")%></div>
<%
end if
%>

<form method="post" name=frmAdd action="upload.asp" enctype="multipart/form-data">
<INPUT type=hidden value=ADD name=action >

  <table width="400" border="1" cellpadding="0" cellspacing="0" bordercolor="#9966cc">
    <tr><td>
		<table width="400" border="0" cellspacing="0" cellpadding="5">
			<tr>
			<td class="trHdr">Attach Files</td>
			<td align="right" class="trHdr" style="WIDTH: 80px">&nbsp;</td>
			</tr>
		</table>
	</td></tr>

	<tr><td>
		<table width="400" border="0" cellspacing="2" cellpadding="5">
			<tr>
			<td width="60" class="lbl1">File</td>
		   <td><input type="file" name="fname1" size=40 value="<%=Request("fname")%>" >
		  </td>
		</tr>
		</table>
	</td></tr>

	<tr height="30">
	<td align="right">
	<input name="cmdAttach" type="button" id="cmd1" value=" Attach " onclick="javascript:submitform();">
	&nbsp;
	</td>
	</tr>

</table>
</html>
