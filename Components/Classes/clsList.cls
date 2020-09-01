VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCode As String

Public Property Let Code(vData As String)
    mstrCode = vData
End Property
Public Property Get Code() As String
    Code = mstrCode
End Property

Public Function RetrieveSeqNo() As String
                                
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='SEQNO'"
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    Set rst = cmd.Execute
        
    If rst.EOF = False Then
        RetrieveSeqNo = Trim(rst.Fields(0))
    End If

    rst.Close
    If RetrieveSeqNo <> "" Then
        cmd.CommandType = adCmdText
        cmd.CommandText = "UPDATE AppValues SET AppData=? WHERE AppCode='SEQNO'"
        cmd.Parameters(0).Value = CLng(RetrieveSeqNo) + 1
        cmd.Execute
        Set rst = Nothing
        Set cmd = Nothing
    Else
        RetrieveSeqNo = "Error"
    End If
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine

End Function

Public Function RetrieveStatus() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveStatus = RetrieveList("SELECT SCode FROM RequestStatus ORDER BY SeqNo")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveDept() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveDept = RetrieveList("SELECT DeptName FROM Dept WHERE Status ='A'")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveAppl() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveAppl = RetrieveList("SELECT ApplName FROM Application WHERE Status ='A'")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveCustomer() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveCustomer = RetrieveList("SELECT CName FROM Customer WHERE Status ='A'")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveCustApplMgr() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveCustApplMgr = RetrieveList("SELECT CName As Customer, ApplName As Application, AppMgr As ITAcctManager FROM ApplicationControl  WHERE Status ='A'")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveApplMgr(strCName As String, strApplName As String) As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveApplMgr = RetrieveList("SELECT AppMgr FROM ApplicationControl WHERE CName=? AND ApplName=? ", _
                                        strCName, strApplName)
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveUser() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveUser = RetrieveList("SELECT UName As UserName, Email, Dept, MgrName, MgrEmail FROM ASRUsers WHERE Status ='A' ")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveITUser() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveITUser = RetrieveList("SELECT UName As UserName FROM ASRUsers WHERE UserType='I' AND Status ='A'")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveUserMgr() As Recordset

On Error GoTo Err_Routine
    
    Set RetrieveUserMgr = RetrieveList("SELECT DISTINCT MgrName, MgrEmail FROM ASRUsers  WHERE Status ='A'")
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveHourRate(Optional ByVal strIP As String) As Integer

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    If strIP = "" Then
        cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='ReqHrRate'"
    ElseIf strIP = "Proj" Then
        cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='ProjHrRate'"
    End If
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    Set rst = cmd.Execute
        
    If rst.EOF = False Then
      RetrieveHourRate = Trim(rst.Fields(0))
    End If

    rst.Close
    
Exit_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveDate() As Date

On Error GoTo Err_Routine
    
    RetrieveDate = ServerDate
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveFormatDate(dt As Date) As String

On Error GoTo Err_Routine
    
    If dt = "12:00:00 AM" Then
        RetrieveFormatDate = ""
    Else
        RetrieveFormatDate = Format(dt, "dd/MM/yyyy")
    End If
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function PopulateDept() As String

On Error GoTo Err_Routine
    Dim rst As ADODB.Recordset
    Dim strValue As String
    
    Set rst = RetrieveDept
    
    Do While rst.EOF = False
        strValue = strValue & "<option value=""" & rst.Fields(0) & """>" & rst.Fields(0) & "</option><br>"
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    PopulateDept = strValue
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set rst = Nothing
    PopulateDept = ""
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveCostEntryDtFrom() As Integer

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='ProjDtFr'"
     
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    Set rst = cmd.Execute
        
    If rst.EOF = False Then
      RetrieveCostEntryDtFrom = Trim(rst.Fields(0))
    End If

    rst.Close
    
Exit_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveCostEntryDtTo() As Integer

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='ProjDtTo'"
     
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    Set rst = cmd.Execute
        
    If rst.EOF = False Then
      RetrieveCostEntryDtTo = Trim(rst.Fields(0))
    End If

    rst.Close
    
Exit_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveCostEntryDtFromMgr() As Integer

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='ProjDtFr1'"
     
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    Set rst = cmd.Execute
        
    If rst.EOF = False Then
      RetrieveCostEntryDtFromMgr = Trim(rst.Fields(0))
    End If

    rst.Close
    
Exit_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveCostEntryDtToMgr() As Integer

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT AppData FROM AppValues WHERE AppCode='ProjDtTo1'"
     
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    Set rst = cmd.Execute
        
    If rst.EOF = False Then
      RetrieveCostEntryDtToMgr = Trim(rst.Fields(0))
    End If

    rst.Close
    
Exit_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

