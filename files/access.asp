<%

Dim objUser, access, intrefno
Set objUser = Server.CreateObject("ASRMaster.clsUser")if Request("refno")="" then 	intrefno=0else	intrefno=Request("refno")end if
access = objUser.RetrieveUserAccess("A1",strUser,intrefno)
 
Set objUser = Nothing

%>
