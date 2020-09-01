<%

Dim objUser, rsUser
Dim appUserType, appAccessLevel
'strUser = "hravi"
Set objUser = Server.CreateObject("ASRMaster.clsUser")
Set rsUser = objUser.RetrieveUserType(strUser)

appUserType = ""
appAccessLevel = ""

If rsUser.EOF = False Then
	appUserType = rsUser.fields("UT")
	appAccessLevel = rsUser.fields("AL")	
End If

Set rsUser = Nothing
Set objUser = Nothing

%>
