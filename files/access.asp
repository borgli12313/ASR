<%

Dim objUser, access, intrefno
Set objUser = Server.CreateObject("ASRMaster.clsUser")
access = objUser.RetrieveUserAccess("A1",strUser,intrefno)
 
Set objUser = Nothing

%>