VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCName       As String
Private mstrApplName    As String
Private mstrAppMgr      As String
Private mstrAppBkpMgr   As String
Private mstrStatus      As String
Private mstrCrUser      As String
Private mdteCrDate      As Date
Private mstrModUser     As String
Private mdteModDate     As Date

'CName, ApplName, AppMgr, AppBkpMgr, CrUser, CrDate, ModUser, ModDate
Private Const mstrInsert As String = _
    "INSERT INTO dbo.ApplicationControl (CName, ApplName, AppMgr, AppBkpMgr, Status, CrUser) " & _
    "VALUES(?, ?, ?, ?, ?, ?) "
Private Const mstrUpdate As String = _
    "UPDATE dbo.ApplicationControl SET AppMgr=?, AppBkpMgr=?, Status=?, ModUser=?, ModDate=getdate() " & _
    "WHERE CName=? AND ApplName=?"
Private Const mstrDelete As String = _
    "DELETE dbo.ApplicationControl WHERE CName = ? AND ApplName=?"
Private Const mstrSelect As String = _
    "SELECT CName, ApplName, AppMgr, AppBkpMgr, Status, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.ApplicationControl " & _
    "WHERE CName = ? AND ApplName=?"
Private Const mstrSelectList As String = _
    "SELECT CName, ApplName, AppMgr, AppBkpMgr, Status, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.ApplicationControl WHERE 1=1 "

Public Property Let CName(vData As String)
    mstrCName = vData
End Property
Public Property Get CName() As String
    CName = mstrCName
End Property

Public Property Let ApplName(vData As String)
    mstrApplName = vData
End Property
Public Property Get ApplName() As String
    ApplName = mstrApplName
End Property

Public Property Let AppMgr(vData As String)
    mstrAppMgr = vData
End Property
Public Property Get AppMgr() As String
    AppMgr = mstrAppMgr
End Property

Public Property Let AppBkpMgr(vData As String)
    mstrAppBkpMgr = vData
End Property
Public Property Get AppBkpMgr() As String
    AppBkpMgr = mstrAppBkpMgr
End Property
Public Property Let CrUser(vData As String)
    mstrCrUser = vData
End Property
Public Property Get CrUser() As String
    CrUser = mstrCrUser
End Property

Public Property Let Status(vData As String)
    mstrStatus = vData
End Property
Public Property Get Status() As String
    Status = mstrStatus
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
    
    blnExists = Me.AMExists(CName, ApplName)
    If blnExists = False Then
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = CnnStrSQL
        cmd.CommandType = adCmdText
        cmd.CommandText = mstrInsert
        cmd.Parameters(0).Value = CName
        cmd.Parameters(1).Value = ApplName
        cmd.Parameters(2).Value = AppMgr
        cmd.Parameters(3).Value = AppBkpMgr
        cmd.Parameters(4).Value = Status
        cmd.Parameters(5).Value = CrUser
        cmd.Execute
        InsertRecord = "OK"
        Set cmd = Nothing
    Else
        InsertRecord = "Exists"
    End If
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
    cmd.Parameters(0).Value = AppMgr
    cmd.Parameters(1).Value = AppBkpMgr
    cmd.Parameters(2).Value = Status
    cmd.Parameters(3).Value = ModUser
    cmd.Parameters(4).Value = CName
    cmd.Parameters(5).Value = ApplName
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

Public Function DeleteRecord(ByVal strCName As String, _
                             ByVal strApplName As String) As String
On Error GoTo Err_Routine
    Dim strResult As String
    strResult = CanDelete(strCName, strApplName)
    
    If strResult = "OK" Then
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = CnnStrSQL
        cmd.CommandType = adCmdText
        cmd.CommandText = mstrDelete
        cmd.Parameters(0).Value = strCName
        cmd.Parameters(1).Value = strApplName
        cmd.Execute
        DeleteRecord = "OK"
        Set cmd = Nothing
    Else
        DeleteRecord = "Exists"
    End If
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    DeleteRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RetrieveSearch(ByVal strCName As String, _
                               ByVal strApplName As String) As Recordset
On Error GoTo Err_Routine
 
    Dim cmd As ADODB.Command
    Dim strSQL As String
    
    strSQL = mstrSelectList
    
    If strCName <> "" Then strSQL = strSQL & " AND CName = '" & strCName & "' "
    If strApplName <> "" Then strSQL = strSQL & " AND ApplName = '" & strApplName & "' "
    
    strSQL = strSQL & " ORDER BY CName "
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    
    Set RetrieveSearch = cmd.Execute
 
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub RetrieveRecord(ByVal strCName As String, _
                          ByVal strApplName As String)
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
                
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelect
    cmd.Parameters(0).Value = strCName
    cmd.Parameters(1).Value = strApplName
    '"SELECT CName, ApplName, AppMgr, AppBkpMgr, Status, CrUser, CrDate, ModUser, ModDate "
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    CName = rst.Fields("CName").Value
    ApplName = rst.Fields("ApplName").Value
    AppMgr = rst.Fields("AppMgr").Value
    AppBkpMgr = rst.Fields("AppBkpMgr").Value
    Status = rst.Fields("Status").Value
    CrUser = ReplaceNull(rst.Fields("CrUser").Value, "")
    CrDate = rst.Fields("CrDate").Value
    ModUser = ReplaceNull(rst.Fields("ModUser").Value, "")
    ModDate = ReplaceNull(rst.Fields("ModDate").Value, "0")
    
    rst.Close
    Set rst = Nothing

    
Exit_Routine:
    Exit Sub
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SetValues(reqObj As Request)
On Error GoTo Err_Routine

    CName = ReplaceNull(Trim(reqObj.Form("cname")), "")
    ApplName = ReplaceNull(Trim(reqObj.Form("applname")), "")
    AppMgr = ReplaceNull(Trim(reqObj.Form("appmgr")), "")
    AppBkpMgr = ReplaceNull(Trim(reqObj.Form("appbkpmgr")), "")
    Status = ReplaceNull(Trim(reqObj.Form("status")), "")
    CrUser = ReplaceNull(Trim(reqObj.Form("uid")), "")
    ModUser = ReplaceNull(Trim(reqObj.Form("uid")), "")

    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function AMExists(ByVal strCName As String, _
                         ByVal strApplName As String) As Boolean

On Error GoTo Err_Routine
    Dim strSQL As String
    
    strSQL = "SELECT CName FROM dbo.ApplicationControl " & _
             "WHERE CName = ? AND ApplName=?"
    
    AMExists = RecordExists(strSQL, strCName, strApplName)

    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CanDelete(ByVal strCName As String, _
                          ByVal strApplName As String) As String

On Error GoTo Err_Routine
    Dim strSQL As String
    Dim blnExist As Boolean
    
   
    strSQL = "SELECT CName FROM dbo.RequestDetails " & _
             "WHERE CName = ? AND ApplName =?"
    
    blnExist = RecordExists(strSQL, strCName, strApplName)
    
    If blnExist = True Then
        CanDelete = "ReqExists"
        GoTo Exit_Routine
    End If
    
    CanDelete = "OK"
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



