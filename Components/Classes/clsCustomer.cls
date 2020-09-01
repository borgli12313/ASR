VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCName       As String
Private mstrDeptName    As String
Private mstrRemarks     As String
Private mstrStatus      As String
Private mstrCType       As String
Private mstrCrUser      As String
Private mdteCrDate      As Date
Private mstrModUser     As String
Private mdteModDate     As Date

Private Const mstrInsert As String = _
    "INSERT INTO dbo.Customer (CName, DeptName, Remarks, Status, CType, CrUser) " & _
    "VALUES(?, ?, ?, ?, ?, ?) "
Private Const mstrUpdate As String = _
    "UPDATE dbo.Customer SET DeptName=?, Remarks=?, Status=?, CType=?, ModUser=?, ModDate=getdate() " & _
    "WHERE CName=?"
Private Const mstrDelete As String = _
    "DELETE dbo.Customer WHERE CName = ?"
Private Const mstrSelect As String = _
    "SELECT CName, DeptName, Remarks, Status, ctype, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.Customer " & _
    "WHERE CName = ?"
Private Const mstrSelectList As String = _
    "SELECT CName, DeptName, Remarks, Status, " & _
    " case ctype when 'N' THEN 'Normal' WHEN 'M' THEN 'Multiclient AC' " & _
    " WHEN 'O' THEN 'Multiclient NonAC' WHEN 'WA' THEN 'Multiclient West A' " & _
    " WHEN 'WD' THEN 'Multiclient West D' WHEN 'MAM' THEN 'Multiclient (AC - MegaHub)' END, " & _
    " CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.Customer WHERE 1=1 "

Public Property Let CName(vData As String)
    mstrCName = vData
End Property
Public Property Get CName() As String
    CName = mstrCName
End Property

Public Property Let DeptName(vData As String)
    mstrDeptName = vData
End Property
Public Property Get DeptName() As String
    DeptName = mstrDeptName
End Property

Public Property Let Remarks(vData As String)
    mstrRemarks = vData
End Property
Public Property Get Remarks() As String
    Remarks = mstrRemarks
End Property

Public Property Let Status(vData As String)
    mstrStatus = vData
End Property
Public Property Get Status() As String
    Status = mstrStatus
End Property

Public Property Let CType(vData As String)
    mstrCType = vData
End Property
Public Property Get CType() As String
    CType = mstrCType
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
    Dim blnResult As Boolean
    blnResult = CustomerExists(CName)
    If blnResult = False Then
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = CnnStrSQL
        cmd.CommandType = adCmdText
        cmd.CommandText = mstrInsert
        cmd.Parameters(0).Value = CName
        cmd.Parameters(1).Value = DeptName
        cmd.Parameters(2).Value = Remarks
        cmd.Parameters(3).Value = Status
        cmd.Parameters(4).Value = CType
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
    cmd.Parameters(0).Value = DeptName
    cmd.Parameters(1).Value = Remarks
    cmd.Parameters(2).Value = Status
    cmd.Parameters(3).Value = CType
    cmd.Parameters(4).Value = ModUser
    cmd.Parameters(5).Value = CName

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

Public Function DeleteRecord(ByVal strCName As String) As String
On Error GoTo Err_Routine
    Dim str As String
    
    str = CanDelete(strCName)
    If str = "OK" Then
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = CnnStrSQL
        cmd.CommandType = adCmdText
        cmd.CommandText = mstrDelete
        cmd.Parameters(0).Value = strCName
        cmd.Execute
        DeleteRecord = "OK"
        Set cmd = Nothing
    Else
        DeleteRecord = str
    End If
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    DeleteRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RetrieveSearch(ByVal strCName As String, _
                               ByVal strDeptName As String, _
                               ByVal strCType As String) As Recordset
On Error GoTo Err_Routine
 
    Dim cmd As ADODB.Command
    Dim strSQL As String
    
    strSQL = mstrSelectList
    
    If strCName <> "" Then strSQL = strSQL & " AND CName LIKE '%" & strCName & "%' "
    If strDeptName <> "" Then strSQL = strSQL & " AND DeptName = '" & strDeptName & "' "
    If strCType <> "" Then strSQL = strSQL & " AND CType = '" & strCType & "' "
    
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

Public Sub RetrieveRecord(ByVal strCName As String)
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelect
    cmd.Parameters(0).Value = strCName
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    CName = rst.Fields("CName").Value
    DeptName = rst.Fields("DeptName").Value
    Remarks = ReplaceNull(rst.Fields("Remarks").Value, "")
    Status = rst.Fields("Status").Value
    CType = rst.Fields("CType").Value
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

Public Sub SetValues(reqObj As Request)
On Error GoTo Err_Routine

    CName = ReplaceNull(Trim(reqObj.Form("cname")), "")
    DeptName = ReplaceNull(Trim(reqObj.Form("dept")), "")
    Remarks = ReplaceNull(Trim(reqObj.Form("remarks")), "")
    Status = ReplaceNull(Trim(reqObj.Form("status")), "")
    CType = ReplaceNull(Trim(reqObj.Form("ctype")), "")
    CrUser = ReplaceNull(Trim(reqObj.Form("uid")), "")
    ModUser = ReplaceNull(Trim(reqObj.Form("uid")), "")

    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CustomerExists(ByVal strCName As String) As Boolean

On Error GoTo Err_Routine
    Dim strSQL As String
    
    strSQL = "SELECT CName FROM dbo.Customer " & _
             "WHERE CName = ?"
    
    CustomerExists = RecordExists(strSQL, strCName)

    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CanDelete(ByVal strCName As String) As String

On Error GoTo Err_Routine
    Dim strSQL As String
    Dim blnExist As Boolean
    
    strSQL = "SELECT CName FROM dbo.ApplicationControl " & _
             "WHERE CName = ?"
    blnExist = RecordExists(strSQL, strCName)
    
    If blnExist = True Then
        CanDelete = "ACExists"
        GoTo Exit_Routine
    End If
    strSQL = "SELECT CName FROM dbo.RequestDetails " & _
             "WHERE CName = ?"
    
    blnExist = RecordExists(strSQL, strCName)
    
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

