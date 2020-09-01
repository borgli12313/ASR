VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDivision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrDivName    As String
Private mstrDivType      As String
Private mstrCrUser      As String
Private mdteCrDate      As Date
Private mstrModUser     As String
Private mdteModDate     As Date
Private mstrMemType     As String

Private Const mstrSelectList As String = _
    "SELECT DivName, case DivType when 'C' then 'Chargable'  " & _
    "when 'N' then 'Not Chargable' else divtype end As DivType, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.Division "

Private Const mstrSelect As String = _
    "SELECT DivName, DivType, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.Division " & _
    "WHERE DivName=? "
    
Private Const mstrInsert As String = _
    "INSERT INTO dbo.Division " & _
    "(DivName, DivType, CrUser) " & _
    "VALUES (?, ?, ?)"
Private Const mstrUpdate As String = _
    "UPDATE dbo.Division SET DivType=?, ModUser=?, ModDate=getdate() " & _
    "WHERE DivName = ?"
Private Const mstrDelete As String = _
    "DELETE dbo.Division " & _
    "WHERE DivName=? "
        
Public Property Let DivName(vData As String)
    mstrDivName = vData
End Property
Public Property Get DivName() As String
    DivName = mstrDivName
End Property

Public Property Let DivType(vData As String)
    mstrDivType = vData
End Property
Public Property Get DivType() As String
    DivType = mstrDivType
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
    blnExists = RecordExists("SELECT DivName FROM dbo.Division WHERE DivName = ?", DivName)
    If blnExists = False Then
        cmd.CommandText = mstrInsert
        cmd.Parameters(0).Value = DivName
        cmd.Parameters(1).Value = DivType
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
    cmd.Parameters(0).Value = DivType
    cmd.Parameters(1).Value = ModUser
    cmd.Parameters(2).Value = Me.DivName

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

Public Sub RetrieveRecord(ByVal strDivName As String)
'retrieve the appl info
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
                
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelect
    cmd.Parameters(0).Value = strDivName
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
     
    DivName = rst.Fields("DivName").Value
    DivType = rst.Fields("DivType").Value
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

Public Function DeleteRecord(ByVal strDivName As String) As String
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim blnExists As Boolean
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    '
    blnExists = RecordExists("SELECT DeptName FROM dbo.Customer WHERE DeptName = ?", strDivName)
    If blnExists = False Then
        cmd.CommandText = mstrDelete
        cmd.Parameters(0).Value = strDivName
        cmd.Execute
        DeleteRecord = "OK"
    Else
        DeleteRecord = "Exists"
    End If
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    DeleteRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function DivExists(ByVal strDivName As String) As String
On Error GoTo Err_Routine

    Dim strSQL As String
    Dim blnExist As Boolean
    
    strSQL = "SELECT DivName FROM dbo.Division WHERE DivName=?"
    
    blnExist = RecordExists(strSQL, strDivName)
    
    If blnExist = True Then
        DivExists = "Exists"
        GoTo Exit_Routine
    End If
        
    DivExists = "OK"
Exit_Routine:
    Exit Function
    
Err_Routine:
    DivExists = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub SetValues(reqObj As Request)
On Error GoTo Err_Routine
 
    DivName = ReplaceNull(Trim(reqObj.Form("divname")), "")
    DivType = ReplaceNull(Trim(reqObj.Form("divtype")), "")
    CrUser = ReplaceNull(Trim(reqObj.Form("uid")), "")
    ModUser = ReplaceNull(Trim(reqObj.Form("uid")), "")
    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


