VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RetrieveUserAccess(ByVal strAccess As String, _
                                    ByVal strUID As String, _
                                    ByVal lngReqNo As Long) As Boolean
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.sASRAccessRights "
    
    cmd.Parameters(1).Value = strAccess
    cmd.Parameters(2).Value = strUID
    cmd.Parameters(3).Value = lngReqNo
    cmd.Parameters(4).Direction = adParamOutput
    cmd.Execute
    RetrieveUserAccess = cmd.Parameters(4).Value
     
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RetrieveUserType(ByVal strUID As String) As Recordset
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT IsNull(UserType,'') As UT , IsNull(AccessLevel,'') As AL FROM dbo.ASRUsers WHERE UID = ?"
    cmd.Parameters(0).Value = strUID
    
    Set RetrieveUserType = cmd.Execute
    
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
