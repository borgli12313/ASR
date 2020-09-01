VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrApplName    As String
Private mstrStatus      As String
Private mstrCrUser      As String
Private mdteCrDate      As Date
Private mstrModUser     As String
Private mdteModDate     As Date
Private mstrMemType     As String

Private Const mstrSelectList As String = _
    "SELECT ApplName, Status, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.Application "

Private Const mstrSelect As String = _
    "SELECT ApplName, Status, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.Application " & _
    "WHERE ApplName=? "
    
Private Const mstrInsert As String = _
    "INSERT INTO dbo.Application " & _
    "(ApplName, Status, CrUser) " & _
    "VALUES (?, ?, ?)"
Private Const mstrUpdate As String = _
    "UPDATE dbo.Application SET Status=?, ModUser=?, ModDate=getdate() " & _
    "WHERE ApplName = ?"
Private Const mstrDelete As String = _
    "DELETE dbo.Application " & _
    "WHERE ApplName=? "
        
Public Property Let ApplName(vData As String)
    mstrApplName = vData
End Property
Public Property Get ApplName() As String
    ApplName = mstrApplName
End Property

Public Property Let Status(vData As String)
    mstrStatus = vData
End Property
Public Property Get Status() As String
    Status = mstrStatus
End Property

Public Property Let CrUser(vData As String)
    mstrCrUser = vData
End Property
Public Property Get CrUser() As String
    CrUser = mstrCrUser
End Property

Public Property Let CrDate(vData As Date)
    mdteCrDate = vData
End Property
Public Property Get CrDate() As Date
    CrDate = mdteCrDate
End Property

Public Property Let ModUser(vData As String)
    mstrModUser = vData
End Property
Public Property Get ModUser() As String
    ModUser = mstrModUser
End Property

Public Property Let ModDate(vData As Date)
    mdteModDate = vData
End Property
Public Property Get ModDate() As Date
    ModDate = mdteModDate
End Property

Public Function InsertRecord() As String
On Error GoTo Err_Routine
    
    Dim blnExists As Boolean
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    blnExists = RecordExists("SELECT ApplName FROM dbo.Application WHERE ApplName = ?", ApplName)
    If blnExists = False Then
        cmd.CommandText = mstrInsert
        cmd.Parameters(0).Value = ApplName
        cmd.Parameters(1).Value = Status
        cmd.Parameters(2).Value = CrUser
        cmd.Execute
        InsertRecord = "OK"
    Else
        InsertRecord = "Exists"
    End If
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    InsertRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function UpdateRecord() As String
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrUpdate
    cmd.Parameters(0).Value = Status
    cmd.Parameters(1).Value = ModUser
    cmd.Parameters(2).Value = Me.ApplName

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
'retrieve the appl info
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

Public Sub RetrieveRecord(ByVal strApplName As String)
'retrieve the appl info
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
                
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelect
    cmd.Parameters(0).Value = strApplName
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
     
    ApplName = rst.Fields("ApplName").Value
    Status = rst.Fields("Status").Value
    CrUser = ReplaceNull(rst.Fields("CrUser").Value, "")
    CrDate = rst.Fields("CrDate").Value
    ModUser = ReplaceNull(rst.Fields("ModUser").Value, "")
    ModDate = ReplaceNull(rst.Fields("ModDate").Value, "0")
    
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

Public Function DeleteRecord(ByVal strApplName As String) As String
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim blnExists As Boolean
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    '
    blnExists = RecordExists("SELECT ApplName FROM dbo.ApplicationControl WHERE ApplName = ?", strApplName)
    If blnExists = False Then
        cmd.CommandText = mstrDelete
        cmd.Parameters(0).Value = strApplName
        cmd.Execute
        DeleteRecord = "OK"
    Else
        DeleteRecord = "ACExists"
    End If
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    DeleteRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ApplExists(ByVal strApplName As String) As String
On Error GoTo Err_Routine

    Dim strSQL As String
    Dim blnExist As Boolean
    
    strSQL = "SELECT ApplName FROM dbo.Application WHERE ApplName=?"
    
    blnExist = RecordExists(strSQL, strApplName)
    
    If blnExist = True Then
        ApplExists = "ACExists"
        GoTo Exit_Routine
    End If
        
    ApplExists = "OK"
Exit_Routine:
    Exit Function
    
Err_Routine:
    ApplExists = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub SetValues(reqObj As Request)
On Error GoTo Err_Routine
 
    ApplName = ReplaceNull(Trim(reqObj.Form("applname")), "")
    Status = ReplaceNull(Trim(reqObj.Form("status")), "")
    CrUser = ReplaceNull(Trim(reqObj.Form("uid")), "")

    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
