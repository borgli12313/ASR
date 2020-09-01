<pre>
<%
'  Variables
'  *********
   Dim objUpload
   Dim intCount, filePath, fileExt, msg
   Dim uploadOk, seqno, mobj

   uploadOk = False
   filePath = "D:\Projects\ASR\upload"  '--PHYSICAL PATH--'
	fileExt = ""
	Set mobj = Server.CreateObject("ASRTrans.clsList")

	seqno=mobj.RetrieveSeqNo


'  Object creation
'  ***************
   Set objUpload = Server.CreateObject("ASPUploadComponent.cUpload")

'  Upload
'  ******
    On Error Resume Next
    objUpload.Form("fname1").SaveFile  filePath, cstr(seqno)+objUpload.Form("fname1").Value, fileExt
    If Err Then
        msg = "Error " & Err.Number & ": " & Err.Description
    Else
        msg = objUpload.Form("fname1").Value & " uploaded.<br>"
        uploadOk = true
    End If

'  Display the number of files uploaded
'  ************************************
   Response.Write(msg)

%>
</pre>

<% if (uploadOk) then %>
<script>
	opener.returnFile("<%= cstr(seqno)+objUpload.Form("fname1").Value %>", "<%= objUpload.Form("fname1").Value %>");
	window.close();
</script>
<% else %>
<script>
	document.location = "attachment.asp?err=1&msg=<%=msg%>&fname=<%= objUpload.Form("fname1").Value %>";
</script>
<%	end if %>
