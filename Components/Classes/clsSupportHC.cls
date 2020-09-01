VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSupportHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCName       As String
Private mstrITStaff     As String
Private mdblHCPercent   As Double
Private mstrRemarks     As String
Private mstrCrUser      As String
Private mdteCrDate      As Date
Private mstrModUser     As String
Private mdteModDate     As Date

'CName, ITStaff, Division, HCPercent, CrUser, CrDate, ModUser, ModDate
Private Const mstrInsert As String = _
    "INSERT INTO dbo.SupportHC (CName, ITStaff, HCPercent, Remarks, CrUser) " & _
    "VALUES(?, ?, ?, ?, ?) "
Private Const mstrUpdate As String = _
    "UPDATE dbo.SupportHC SET HCPercent=?, Remarks=?, ModUser=?, ModDate=getdate() " & _
    "WHERE CName=? AND ITStaff=?"
Private Const mstrDelete As String = _
    "DELETE dbo.SupportHC WHERE CName = ? AND ITStaff=?"
Private Const mstrSelect As String = _
    "SELECT CName, ITStaff, HCPercent, Remarks, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.SupportHC " & _
    "WHERE CName = ? AND ITStaff=?"
Private Const mstrSelectList As String = _
    "SELECT CName, ITStaff, HCPercent, Remarks, CrUser, CrDate, ModUser, ModDate " & _
    "FROM dbo.SupportHC WHERE 1=1 "

Public Property Let CName(vData As String)
    mstrCName = vData
End Property
Public Property Get CName() As String
    CName = mstrCName
End Property

Public Property Let ITStaff(vData As String)
    mstrITStaff = vData
End Property
Public Property Get ITStaff() As String
    ITStaff = mstrITStaff
End Property

Public Property Let HCPercent(vData As Double)
    mdblHCPercent = vData
End Property
Public Property Get HCPercent() As Double
    HCPercent = mdblHCPercent
End Property

Public Property Let CrUser(vData As String)
    mstrCrUser = vData
End Property
Public Property Get CrUser() As String
    CrUser = mstrCrUser
End Property

Public Property Let Remarks(vData As String)
    mstrRemarks = vData
End Property
Public Property Get Remarks() As String
    Remarks = mstrRemarks
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
    
    blnExists = SupportHCExists(CName, ITStaff)
    If blnExists = False Then
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = CnnStrSQL
        cmd.CommandType = adCmdText
        cmd.CommandText = mstrInsert
        'CName, ITStaff, Division, HCPercent, Remarks, CrUser
        cmd.Parameters(0).Value = CName
        cmd.Parameters(1).Value = ITStaff
        cmd.Parameters(2).Value = HCPercent
        cmd.Parameters(3).Value = Remarks
        cmd.Parameters(4).Value = CrUser
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
    cmd.Parameters(0).Value = HCPercent
    cmd.Parameters(1).Value = Remarks
    cmd.Parameters(2).Value = ModUser
    cmd.Parameters(3).Value = CName
    cmd.Parameters(4).Value = ITStaff
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
                             ByVal strITStaff As String) As String
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrDelete
    cmd.Parameters(0).Value = strCName
    cmd.Parameters(1).Value = strITStaff
    cmd.Execute
    DeleteRecord = "OK"
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    DeleteRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RetrieveSearch(ByVal strCName As String, _
                               ByVal strITStaff As String) As Recordset
On Error GoTo Err_Routine
 
    Dim cmd As ADODB.Command
    Dim strSQL As String
    
    strSQL = mstrSelectList
    
    If strCName <> "" Then strSQL = strSQL & " AND CName = '" & strCName & "' "
    If strITStaff <> "" Then strSQL = strSQL & " AND ITStaff = '" & strITStaff & "' "
    
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
                          ByVal strITStaff As String)
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
                
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = mstrSelect
    cmd.Parameters(0).Value = strCName
    cmd.Parameters(1).Value = strITStaff
    '"SELECT CName, ITStaff, Division, HCPercent, Remarks, CrUser, CrDate, ModUser, ModDate "
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    CName = rst.Fields("CName").Value
    ITStaff = rst.Fields("ITStaff").Value
    HCPercent = rst.Fields("HCPercent").Value
    Remarks = ReplaceNull(rst.Fields("Remarks").Value, "")
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
    ITStaff = ReplaceNull(Trim(reqObj.Form("itstaff")), "")
    HCPercent = ReplaceNull(Trim(reqObj.Form("hcpercent")), "")
    Remarks = ReplaceNull(Trim(reqObj.Form("remarks")), "")
    CrUser = ReplaceNull(Trim(reqObj.Form("uid")), "")
    ModUser = ReplaceNull(Trim(reqObj.Form("uid")), "")

    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function SupportHCExists(ByVal strCName As String, _
                                ByVal strITStaff As String) As Boolean

On Error GoTo Err_Routine
    Dim strSQL As String
    
    strSQL = "SELECT CName FROM dbo.SupportHC " & _
             "WHERE CName = ? AND ITStaff=?"
    
    SupportHCExists = RecordExists(strSQL, strCName, strITStaff)

    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CanDelete(ByVal strCName As String, _
                          ByVal strITStaff As String) As String

On Error GoTo Err_Routine
    Dim strSQL As String
    Dim blnExist As Boolean
    
   
    strSQL = "SELECT CName FROM dbo.RequestDetails " & _
             "WHERE CName = ? AND ITStaff =?"
    
    blnExist = RecordExists(strSQL, strCName, strITStaff)
    
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
