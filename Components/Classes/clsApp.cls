VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrAppCode    As String
Private mstrAppData      As String
Private mstrRemarks     As String

Private Const mstrSelectList As String = _
    "SELECT AppCode, AppData, Remarks FROM dbo.AppValues  " & _
    "WHERE AppCode LIKE 'ProjDt%' " & _
    "ORDER By Remarks"
Private Const mstrSelect As String = _
    "SELECT AppCode, AppData, Remarks FROM dbo.AppValues " & _
    "WHERE AppCode=? "
    
Private Const mstrUpdate As String = _
    "UPDATE dbo.AppValues SET AppData=? " & _
    "WHERE AppCode = ?"

Public Property Let AppCode(vData As String)
    mstrAppCode = vData
End Property
Public Property Get AppCode() As String
    AppCode = mstrAppCode
End Property

Public Property Let AppData(vData As String)
    mstrAppData = vData
End Property
Public Property Get AppData() As String
    AppData = mstrAppData
End Property

Public Property Let Remarks(vData As String)
    mstrRemarks = vData
End Property
Public Property Get Remarks() As String
    Remarks = mstrRemarks
End Property

Public Function UpdateRecord() As String
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrUpdate
    cmd.Parameters(0).Value = AppData
    cmd.Parameters(1).Value = AppCode
    cmd.Execute
        
    UpdateRecord = "OK"
    
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    UpdateRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RetrieveList() As Recordset
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
                
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelectList
    
    Set RetrieveList = cmd.Execute
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub RetrieveRecord(ByVal strAppCode As String)
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
                
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelect
    cmd.Parameters(0).Value = strAppCode
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
     
    AppCode = rst.Fields("AppCode").Value
    AppData = rst.Fields("AppData").Value
    Remarks = rst.Fields("Remarks").Value

    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    
Exit_Routine:
    Exit Sub
    
Err_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SetValues(reqObj As Request)
On Error GoTo Err_Routine
 
    AppCode = ReplaceNull(Trim(reqObj.Form("AppCode")), "")
    AppData = ReplaceNull(Trim(reqObj.Form("AppData")), "")

    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

