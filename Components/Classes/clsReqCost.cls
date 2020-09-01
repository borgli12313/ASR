VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReqCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mintReqNo      As Integer
Private mintRCYear      As Integer
Private mintRCMonth    As Integer
Private mintCCHourRate    As Integer
Private mintNCCHourRate    As Integer
Private mintCCManHour As Integer
Private mintNCCManHour As Integer
Private mdblChargeCost    As Double
Private mdblNonChargeCost As Double
Private mstrCrUser      As String
Private mdteCrDate      As Date

Public Property Let ReqNo(vData As Integer)
    mintReqNo = vData
End Property
Public Property Get ReqNo() As Integer
    ReqNo = mintReqNo
End Property

Public Property Let RCYear(vData As Integer)
    mintRCYear = vData
End Property
Public Property Get RCYear() As Integer
    RCYear = mintRCYear
End Property

Public Property Let RCMonth(vData As Integer)
    mintRCMonth = vData
End Property
Public Property Get RCMonth() As Integer
    RCMonth = mintRCMonth
End Property

Public Property Let CCManHour(vData As Integer)
    mintCCManHour = vData
End Property
Public Property Get CCManHour() As Integer
    CCManHour = mintCCManHour
End Property

Public Property Let CCHourRate(vData As Integer)
    mintCCHourRate = vData
End Property
Public Property Get CCHourRate() As Integer
    CCHourRate = mintCCHourRate
End Property

Public Property Let ChargeCost(vData As Double)
    mdblChargeCost = vData
End Property
Public Property Get ChargeCost() As Double
    ChargeCost = mdblChargeCost
End Property

Public Property Let NCCManHour(vData As Integer)
    mintNCCManHour = vData
End Property
Public Property Get NCCManHour() As Integer
    NCCManHour = mintNCCManHour
End Property

Public Property Let NCCHourRate(vData As Integer)
    mintNCCHourRate = vData
End Property
Public Property Get NCCHourRate() As Integer
    NCCHourRate = mintNCCHourRate
End Property

Public Property Let NonChargeCost(vData As Double)
    mdblNonChargeCost = vData
End Property
Public Property Get NonChargeCost() As Double
    NonChargeCost = mdblNonChargeCost
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

Public Function InsertReqCost() As String
On Error GoTo Err_Routine

    Dim cmd      As ADODB.Command
    Dim cnn      As ADODB.Connection
    Dim intTrans As Integer
    
    Set cnn = SQLConnWrite
    cnn.Open
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdText
    
        
    intTrans = cnn.BeginTrans
        cmd.CommandText = "DELETE dbo.RequestCost WHERE ReqNo=? AND  RCYear=? AND RCMonth=?"
        cmd.Parameters(0).Value = ReqNo
        cmd.Parameters(1).Value = RCYear
        cmd.Parameters(2).Value = RCMonth
        cmd.Execute
        cmd.CommandText = "INSERT INTO dbo.RequestCost " & _
                      "(ReqNo, RCYear, RCMonth, CCManHour, CCHourRate, ChargeCost, " & _
                      " NCCManHour, NCCHourRate, NonChargeCost, CrUser ) " & _
                      "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        cmd.Parameters(0).Value = ReqNo
        cmd.Parameters(1).Value = RCYear
        cmd.Parameters(2).Value = RCMonth
        cmd.Parameters(3).Value = CCManHour
        cmd.Parameters(4).Value = CCHourRate
        cmd.Parameters(5).Value = ChargeCost
        cmd.Parameters(6).Value = NCCManHour
        cmd.Parameters(7).Value = NCCHourRate
        cmd.Parameters(8).Value = NonChargeCost
        cmd.Parameters(9).Value = CrUser
        cmd.Execute
    cnn.CommitTrans
      
    InsertReqCost = "OK"
    
    Set cmd = Nothing
    cnn.Close
    Set cnn = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    If intTrans = 1 Then
        cnn.RollbackTrans
        intTrans = intTrans - 1
    End If
    If Not cnn Is Nothing Then
        If cnn.State = adStateOpen Then cnn.Close
        Set cnn = Nothing
    End If
    InsertReqCost = ""
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub RetrieveRequestCost()

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT ReqNo, RCYear, RCMonth, CCManHour, CCHourRate, ChargeCost, " & _
                      " NCCManHour, NCCHourRate, NonChargeCost, CrUser , CrDate " & _
                        "FROM dbo.RequestCost WHERE ReqNo = ? AND  RCYear=? AND RCMonth=?"
                        
    cmd.Parameters(0).Value = ReqNo
    cmd.Parameters(1).Value = RCYear
    cmd.Parameters(2).Value = RCMonth
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    If rst.EOF = False Then
        'ReqNo = rst.Fields("ReqNo").Value
        'RCYear = rst.Fields("RCYear").Value
        'RCMonth = rst.Fields("RCMonth").Value
        CCManHour = rst.Fields("CCManHour").Value
        CCHourRate = rst.Fields("CCHourRate").Value
        ChargeCost = rst.Fields("ChargeCost").Value
        NCCManHour = rst.Fields("NCCManHour").Value
        NCCHourRate = rst.Fields("NCCHourRate").Value
        NonChargeCost = rst.Fields("NonChargeCost").Value
        CrUser = rst.Fields("CrUser").Value
        CrDate = rst.Fields("CrDate").Value
    Else
        Dim cls As clsList
        Dim intRate As Integer
        
        Set cls = New clsList
        
        intRate = cls.RetrieveHourRate("Proj")
        CCHourRate = intRate
        NCCHourRate = intRate
        
        Set cls = Nothing
    End If
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    
Exit_Routine:
    Exit Sub
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Sub

Public Function RetrieveRequestCostList() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT ReqNo, RCYear, RCMonth, ChargeCost, NonChargeCost, CrUser , CrDate " & _
                        "FROM dbo.RequestCost WHERE ReqNo = ? "
                        
    cmd.Parameters(0).Value = ReqNo
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveRequestCostList = rst
    
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Sub SetValues(reqObj As Request)
On Error GoTo Err_Routine

    ReqNo = Trim(reqObj.Form("ReqNo"))
    RCYear = Trim(reqObj.Form("rcyear"))
    RCMonth = ReplaceNull(Trim(reqObj.Form("rcmonth")), "")
    ChargeCost = ReplaceNull(Trim(reqObj.Form("chargecost")), "")
    CCManHour = ReplaceNull(Trim(reqObj.Form("ccmanhour")), "")
    CCHourRate = ReplaceNull(Trim(reqObj.Form("cchourrate")), "")
    NCCManHour = ReplaceNull(Trim(reqObj.Form("nccmanhour")), "")
    NCCHourRate = ReplaceNull(Trim(reqObj.Form("ncchourrate")), "")
    NonChargeCost = ReplaceNull(Trim(reqObj.Form("nonchargecost")), "")
    CrUser = ReplaceNull(Trim(reqObj.Form("UID")), "")
        
Exit_Routine:
    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


