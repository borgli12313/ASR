VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequestDets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mintReqNo      As Integer
Private mstrSCode      As String
Private mstrReqUser    As String
Private mstrReqEmail    As String
Private mstrReqUserMgr  As String
Private mstrReqMgrEmail As String
Private mstrCName      As String
Private mstrDeptName   As String
Private mstrApplName   As String
Private mstrAppMgr     As String
Private mintPriority   As Integer
Private mstrReqTitle   As String
Private mstrReqDesc    As String
Private mdteExpCloseDate   As Date
Private mstrExpCloseDateStr   As String
Private mstrRemarksCH  As String
Private mstrTeamLead   As String
Private mstrDeveloper  As String
Private mstrExpQC      As String
Private mintEstManHour As Integer
Private mintEstHourRate As Integer
Private mdblEstTotalCost   As Double
Private mdteExpStartDate   As Date
Private mstrExpStartDateStr   As String
Private mdteExpEndDate As Date
Private mstrExpEndDateStr   As String
Private mstrRemarksIT  As String
Private mdteActStartDate   As Date
Private mstrActStartDateStr   As String
Private mstrRemarksDev     As String
Private mintActManHour As Integer
Private mintActHourRate As Integer
Private mdblActTotalCost   As Double
Private mdteActEndDate As Date
Private mstrActEndDateStr   As String
Private mstrActQC      As String
Private mstrStatusQC   As String
Private mstrRemarksQC  As String
Private mdteUATReadyDate   As Date
Private mstrUATReadyDateStr   As String
Private mstrUserUAT    As String
Private mstrStatusUAT  As String
Private mstrRemarksUAT      As String
Private mstrIsUATFailed     As String
Private mintNoUATFailed As Integer
Private mdteExpCutinDate    As Date
Private mstrExpCutinDateStr As String
Private mstrDeployUser      As String
Private mdteActCutinDate    As Date
Private mstrActCutinDateStr As String
Private mstrRemarksDeploy   As String
Private mstrReqType   As String
Private mstrCrUser      As String
Private mdteCrDate      As Date
Private mstrModUser     As String
Private mdteModDate     As Date
Private marrFileName()
Private marrFileNameDisplay()
Private marrFileUser()
Private marrFileDate()

Public Property Let ReqNo(vData As Integer)
    mintReqNo = vData
End Property
Public Property Get ReqNo() As Integer
    ReqNo = mintReqNo
End Property

Public Property Let SCode(vData As String)
    mstrSCode = vData
End Property
Public Property Get SCode() As String
    SCode = mstrSCode
End Property

Public Property Let ReqUser(vData As String)
    mstrReqUser = vData
End Property
Public Property Get ReqUser() As String
    ReqUser = mstrReqUser
End Property

Public Property Let ReqEmail(vData As String)
    mstrReqEmail = vData
End Property
Public Property Get ReqEmail() As String
    ReqEmail = mstrReqEmail
End Property

Public Property Let ReqUserMgr(vData As String)
    mstrReqUserMgr = vData
End Property
Public Property Get ReqUserMgr() As String
    ReqUserMgr = mstrReqUserMgr
End Property

Public Property Let ReqMgrEmail(vData As String)
    mstrReqMgrEmail = vData
End Property
Public Property Get ReqMgrEmail() As String
    ReqMgrEmail = mstrReqMgrEmail
End Property

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

Public Property Let Priority(vData As Integer)
    mintPriority = vData
End Property
Public Property Get Priority() As Integer
    Priority = mintPriority
End Property

Public Property Let ReqTitle(vData As String)
    mstrReqTitle = vData
End Property
Public Property Get ReqTitle() As String
    ReqTitle = mstrReqTitle
End Property

Public Property Let ReqDesc(vData As String)
    mstrReqDesc = vData
End Property
Public Property Get ReqDesc() As String
    ReqDesc = mstrReqDesc
End Property

Public Property Let ExpCloseDate(vData As Date)
    mdteExpCloseDate = vData
End Property
Public Property Get ExpCloseDate() As Date
    ExpCloseDate = mdteExpCloseDate
End Property

Public Property Let ExpCloseDateStr(vData As String)
    mstrExpCloseDateStr = vData
End Property
Public Property Get ExpCloseDateStr() As String
    ExpCloseDateStr = mstrExpCloseDateStr
End Property

Public Property Let RemarksCH(vData As String)
    mstrRemarksCH = vData
End Property
Public Property Get RemarksCH() As String
    RemarksCH = mstrRemarksCH
End Property

Public Property Let TeamLead(vData As String)
    mstrTeamLead = vData
End Property
Public Property Get TeamLead() As String
    TeamLead = mstrTeamLead
End Property

Public Property Let Developer(vData As String)
    mstrDeveloper = vData
End Property
Public Property Get Developer() As String
    Developer = mstrDeveloper
End Property

Public Property Let ExpQC(vData As String)
    mstrExpQC = vData
End Property
Public Property Get ExpQC() As String
    ExpQC = mstrExpQC
End Property

Public Property Let EstManHour(vData As Integer)
    mintEstManHour = vData
End Property
Public Property Get EstManHour() As Integer
    EstManHour = mintEstManHour
End Property

Public Property Let EstHourRate(vData As Integer)
    mintEstHourRate = vData
End Property
Public Property Get EstHourRate() As Integer
    EstHourRate = mintEstHourRate
End Property

Public Property Let EstTotalCost(vData As Double)
    mdblEstTotalCost = vData
End Property
Public Property Get EstTotalCost() As Double
    EstTotalCost = mdblEstTotalCost
End Property

Public Property Let ExpStartDate(vData As Date)
    mdteExpStartDate = vData
End Property
Public Property Get ExpStartDate() As Date
    ExpStartDate = mdteExpStartDate
End Property

Public Property Let ExpStartDateStr(vData As String)
    mstrExpStartDateStr = vData
End Property
Public Property Get ExpStartDateStr() As String
    ExpStartDateStr = mstrExpStartDateStr
End Property

Public Property Let ExpEndDate(vData As Date)
    mdteExpEndDate = vData
End Property
Public Property Get ExpEndDate() As Date
    ExpEndDate = mdteExpEndDate
End Property

Public Property Let ExpEndDateStr(vData As String)
    mstrExpEndDateStr = vData
End Property
Public Property Get ExpEndDateStr() As String
    ExpEndDateStr = mstrExpEndDateStr
End Property

Public Property Let RemarksIT(vData As String)
    mstrRemarksIT = vData
End Property
Public Property Get RemarksIT() As String
    RemarksIT = mstrRemarksIT
End Property

Public Property Let ActStartDate(vData As Date)
    mdteActStartDate = vData
End Property
Public Property Get ActStartDate() As Date
    ActStartDate = mdteActStartDate
End Property

Public Property Let ActStartDateStr(vData As String)
    mstrActStartDateStr = vData
End Property
Public Property Get ActStartDateStr() As String
    ActStartDateStr = mstrActStartDateStr
End Property

Public Property Let RemarksDev(vData As String)
    mstrRemarksDev = vData
End Property
Public Property Get RemarksDev() As String
    RemarksDev = mstrRemarksDev
End Property

Public Property Let ActManHour(vData As Integer)
    mintActManHour = vData
End Property
Public Property Get ActManHour() As Integer
    ActManHour = mintActManHour
End Property

Public Property Let ActHourRate(vData As Integer)
    mintActHourRate = vData
End Property
Public Property Get ActHourRate() As Integer
    ActHourRate = mintActHourRate
End Property

Public Property Let ActTotalCost(vData As Double)
    mdblActTotalCost = vData
End Property
Public Property Get ActTotalCost() As Double
    ActTotalCost = mdblActTotalCost
End Property

Public Property Let ActEndDate(vData As Date)
    mdteActEndDate = vData
End Property
Public Property Get ActEndDate() As Date
    ActEndDate = mdteActEndDate
End Property

Public Property Let ActEndDateStr(vData As String)
    mstrActEndDateStr = vData
End Property
Public Property Get ActEndDateStr() As String
    ActEndDateStr = mstrActEndDateStr
End Property

Public Property Let ActQC(vData As String)
    mstrActQC = vData
End Property
Public Property Get ActQC() As String
    ActQC = mstrActQC
End Property

Public Property Let StatusQC(vData As String)
    mstrStatusQC = vData
End Property
Public Property Get StatusQC() As String
    StatusQC = mstrStatusQC
End Property

Public Property Let RemarksQC(vData As String)
    mstrRemarksQC = vData
End Property
Public Property Get RemarksQC() As String
    RemarksQC = mstrRemarksQC
End Property

Public Property Let UATReadyDate(vData As Date)
    mdteUATReadyDate = vData
End Property
Public Property Get UATReadyDate() As Date
    UATReadyDate = mdteUATReadyDate
End Property

Public Property Let UATReadyDateStr(vData As String)
    mstrUATReadyDateStr = vData
End Property
Public Property Get UATReadyDateStr() As String
    UATReadyDateStr = mstrUATReadyDateStr
End Property

Public Property Let UserUAT(vData As String)
    mstrUserUAT = vData
End Property
Public Property Get UserUAT() As String
    UserUAT = mstrUserUAT
End Property

Public Property Let StatusUAT(vData As String)
    mstrStatusUAT = vData
End Property
Public Property Get StatusUAT() As String
    StatusUAT = mstrStatusUAT
End Property

Public Property Let RemarksUAT(vData As String)
    mstrRemarksUAT = vData
End Property
Public Property Get RemarksUAT() As String
    RemarksUAT = mstrRemarksUAT
End Property

Public Property Let ExpCutinDate(vData As Date)
    mdteExpCutinDate = vData
End Property
Public Property Get ExpCutinDate() As Date
    ExpCutinDate = mdteExpCutinDate
End Property

Public Property Let ExpCutinDateStr(vData As String)
    mstrExpCutinDateStr = vData
End Property
Public Property Get ExpCutinDateStr() As String
    ExpCutinDateStr = mstrExpCutinDateStr
End Property

Public Property Let IsUATFailed(vData As String)
    mstrIsUATFailed = vData
End Property
Public Property Get IsUATFailed() As String
    IsUATFailed = mstrIsUATFailed
End Property

Public Property Let NoUATFailed(vData As Integer)
    mintNoUATFailed = vData
End Property
Public Property Get NoUATFailed() As Integer
    NoUATFailed = mintNoUATFailed
End Property

Public Property Let DeployUser(vData As String)
    mstrDeployUser = vData
End Property
Public Property Get DeployUser() As String
    DeployUser = mstrDeployUser
End Property

Public Property Let ActCutinDate(vData As Date)
    mdteActCutinDate = vData
End Property
Public Property Get ActCutinDate() As Date
    ActCutinDate = mdteActCutinDate
End Property

Public Property Let ActCutinDateStr(vData As String)
    mstrActCutinDateStr = vData
End Property
Public Property Get ActCutinDateStr() As String
    ActCutinDateStr = mstrActCutinDateStr
End Property

Public Property Let RemarksDeploy(vData As String)
    mstrRemarksDeploy = vData
End Property
Public Property Get RemarksDeploy() As String
    RemarksDeploy = mstrRemarksDeploy
End Property

Public Property Let ReqType(vData As String)
    mstrReqType = vData
End Property
Public Property Get ReqType() As String
    ReqType = mstrReqType
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

    Dim cmd      As ADODB.Command
    Dim cnn      As ADODB.Connection
    Dim rst      As ADODB.Recordset
    Dim intTrans As Integer
    Dim strTo          As String
    Dim strCC          As String
    Dim strBody        As String
    Dim strSubject     As String
    Dim strTemp        As String

    Set cnn = SQLConnWrite
    cnn.Open
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnn
    intTrans = cnn.BeginTrans
       '(ReqUser, UserEmail, ReqUserMgr, UserMgrEmail, 5
        'CName, DeptName, ApplName, AppMgr,9
        'Priority , ReqTitle, ReqDesc, ExpCloseDate, CrUser 14
    
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "sRequestDetsIns"
        
        cmd.Parameters(1).Value = ReqUser
        cmd.Parameters(2).Value = ReqEmail
        cmd.Parameters(3).Value = ReqUserMgr
        cmd.Parameters(4).Value = ReqMgrEmail
        cmd.Parameters(5).Value = CName
        cmd.Parameters(6).Value = DeptName
        cmd.Parameters(7).Value = ApplName
        cmd.Parameters(8).Value = AppMgr
        cmd.Parameters(9).Value = Priority
        cmd.Parameters(10).Value = ReqTitle
        cmd.Parameters(11).Value = ReqDesc
        cmd.Parameters(12).Value = ExpCloseDate '"12/02/2003"
        cmd.Parameters(13).Value = CrUser
        cmd.Parameters(14).Value = ReqType
        cmd.Parameters(15).Direction = adParamOutput
        cmd.Execute
        ReqNo = cmd.Parameters(15).Value
        
        cmd.CommandType = adCmdText
        cmd.CommandText = "INSERT INTO RequestDetailsStatus (ReqNo, SCode, CrUser) " & _
                          "VALUES (?, ?, ?)"
        cmd.Parameters(0).Value = ReqNo
        cmd.Parameters(1).Value = "Open"
        cmd.Parameters(2).Value = CrUser
        cmd.Execute
        
        Dim intarr As Integer
        Dim i      As Integer
    
        intarr = UBound(marrFileName)
        
        cmd.CommandText = " INSERT INTO RequestAttach " & _
                        "(ReqNo, DocRefNo, DocFileName, CrUser,CrDate) " & _
                        "VALUES (?,?,?,?,?) "
        For i = 1 To intarr
            cmd.Parameters(0).Value = ReqNo
            cmd.Parameters(1).Value = marrFileName(i)
            cmd.Parameters(2).Value = marrFileNameDisplay(i)
            cmd.Parameters(3).Value = marrFileUser(i)
            cmd.Parameters(4).Value = marrFileDate(i)
            cmd.Execute
        Next
        
        'insert user if not exists
        cmd.CommandText = " SELECT UName FROM ASRUsers WHERE UName=? "
        cmd.Parameters(0).Value = ReqUser
        
        Set rst = New ADODB.Recordset
        rst.ActiveConnection = CnnStrSQL
        rst.CursorType = adOpenStatic
        rst.LockType = adLockReadOnly
        rst.CursorLocation = adUseClient
        rst.Open cmd
        
        If rst.EOF = True Then
            cmd.CommandText = "INSERT INTO ASRUsers([UName], [Email], [Dept], [MgrName], [MgrEmail], CrUser) " & _
                              "VALUES (?,?,?,?,?,?) "
            cmd.Parameters(0).Value = ReqUser
            cmd.Parameters(1).Value = ReqEmail
            cmd.Parameters(2).Value = DeptName
            cmd.Parameters(3).Value = ReqUserMgr
            cmd.Parameters(4).Value = ReqMgrEmail
            cmd.Parameters(5).Value = CrUser
        Else
            cmd.CommandText = "UPDATE ASRUsers SET Email=?, Dept=?, MgrName=?, MgrEmail=?, ModUser=? " & _
                              " WHERE UName=? "
            cmd.Parameters(0).Value = ReqEmail
            cmd.Parameters(1).Value = DeptName
            cmd.Parameters(2).Value = ReqUserMgr
            cmd.Parameters(3).Value = ReqMgrEmail
            cmd.Parameters(4).Value = CrUser
            cmd.Parameters(5).Value = ReqUser

        End If
        cmd.Execute
        rst.Close
    cnn.CommitTrans
    
On Error GoTo ErrMail_Routine
    cmd.CommandText = "SELECT A.Email, B.Email  FROM ApplicationControl C " & _
                        "INNER join dbo.ASRUsers A on A.UName=C.AppMgr " & _
                        "INNER join dbo.ASRUsers B on B.UName=C.AppBkpMgr " & _
                        "WHERE C.CName=? AND C.ApplName=? "
    cmd.Parameters(0).Value = CName
    cmd.Parameters(1).Value = ApplName
    rst.Open cmd
    If rst.EOF = False Then
        strTo = rst.Fields(0).Value & ", "
        strCC = rst.Fields(1).Value & ", "
    End If
    strCC = strCC & ReqEmail & ", " & ReqMgrEmail & ", "
    rst.Close
    
    cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName = (SELECT AppData FROM AppValues WHERE AppCode='SupMgr')"
    rst.Open cmd
    If rst.EOF = False Then strCC = strCC & rst.Fields(0).Value & ", "
    rst.Close
    
    strSubject = "ASR New Request - ReqNo: " & ReqNo & " - " & ApplName
    
    strBody = "REQUEST NO : " & ReqNo & vbCrLf
    strBody = strBody & "REQUESTOR : " & ReqUser & vbCrLf
    strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
    strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
    strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
    strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
    strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf
    strBody = strBody & "EXP. GO LIVE DATE : " & Me.ExpCloseDate & vbCrLf & vbCrLf
    strBody = strBody & "Please login to the ASR Application for the details " & vbCrLf
    If strTo <> "" Then
        strTemp = SendMail(strTo, strCC, strSubject, strBody)
        If strTemp = "1" Then
            
        Else
            GoTo ErrMail_Routine
        End If
    End If
    
    InsertRecord = "OK"
    
    Set cmd = Nothing
    cnn.Close
    Set cnn = Nothing
    
    GoTo Exit_Routine
    
ErrMail_Routine:
     InsertRecord = "ErrMail"
     Set cmd = Nothing
     cnn.Close
     Set cnn = Nothing
     Err.Raise Err.Number, Err.Source, Err.Description
     GoTo Exit_Routine
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
    InsertRecord = ""
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function UpdateRecord() As String

On Error GoTo Err_Routine
    
    Dim cnn      As ADODB.Connection
    Dim cmd      As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim intTrans As Integer
    Dim intCS As Integer
    Dim strTo          As String
    Dim strCC          As String
    Dim strBody        As String
    Dim strSubject     As String
    Dim strTemp        As String
    Dim strdev         As String
    Dim dteActEndDateOrg As Date
    
    Dim strTeamLeadOrg      As String
    Dim strDeveloperOrg     As String
    Dim strExpQCOrg         As String
    Dim intEstManHourOrg    As Integer
    Dim dblEstTotalCostOrg  As Double
    Dim dteExpStartDateOrg  As Date
    Dim dteExpEndDateOrg    As Date

    Set cnn = SQLConnWrite
    cnn.Open
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdText
    intCS = RetrieveStatusNo(SCode)
    
    intTrans = cnn.BeginTrans
        cmd.CommandText = "SELECT isnull(ActEndDate,''), isnull(TeamLead,''), isnull(Developer,''), " & _
                          "isnull(ExpQC,''), isnull(EstManHour,0), isnull(EstTotalCost,0), " & _
                          "isnull(ExpStartDate,''), isnull(ExpEndDate,'') " & _
                          "FROM RequestDetails " & _
                          "WHERE ReqNo =?"
        cmd.Parameters(0).Value = ReqNo
        Set rst = New ADODB.Recordset
        rst.ActiveConnection = CnnStrSQL
        rst.CursorType = adOpenStatic
        rst.LockType = adLockReadOnly
        rst.CursorLocation = adUseClient
        rst.Open cmd

        If rst.EOF = False Then
            dteActEndDateOrg = rst.Fields(0).Value
            strTeamLeadOrg = rst.Fields(1).Value
            strDeveloperOrg = rst.Fields(2).Value
            If strDeveloperOrg <> "" And InStr(1, strDeveloperOrg, ",") > 0 Then strDeveloperOrg = Mid(strDeveloperOrg, 1, InStr(1, strDeveloperOrg, ",") - 1)
            strExpQCOrg = rst.Fields(3).Value
            intEstManHourOrg = rst.Fields(4).Value
            dblEstTotalCostOrg = rst.Fields(5).Value
            dteExpStartDateOrg = Format(rst.Fields(6).Value, "dd/MM/yyyy")
            dteExpEndDateOrg = Format(rst.Fields(7).Value, "dd/MM/yyyy")
        End If
        rst.Close
        cmd.CommandText = "UPDATE RequestDetails " & _
                          "SET ReqUser=?, UserEmail=?, ReqUserMgr=?, UserMgrEmail=?, CName =?, DeptName =?, " & _
                          "ApplName =?, AppMgr =?, Priority=?, ReqTitle =?, ReqDesc =?, " & _
                          "ExpCloseDate=?, TeamLead=?, Developer=?, ExpQC=?, EstManHour=?, " & _
                          "EstTotalCost=?, ExpStartDate=?, ExpEndDate=?, RemarksIT =?, " & _
                          "ActStartDate=?, RemarksDev =?, ActManHour=?, ActTotalCost=?, " & _
                          "ActEndDate=?, ActQC=?, StatusQC=?, RemarksQC =?, UATReadyDate=?, UserUAT=?, " & _
                          "StatusUAT=?, RemarksUAT =?, ExpCutinDate=?, DeployUser=?, " & _
                          "ActCutinDate=?, RemarksDeploy =?, IsUATFailed=?, ModUser =?, ModDate=getdate(), " & _
                          "EstHourRate=?, ActHourRate=? WHERE ReqNo =?"

        cmd.Parameters(0).Value = ReqUser
        cmd.Parameters(1).Value = ReqEmail
        cmd.Parameters(2).Value = ReqUserMgr
        cmd.Parameters(3).Value = ReqMgrEmail
        cmd.Parameters(4).Value = CName
        cmd.Parameters(5).Value = DeptName
        cmd.Parameters(6).Value = ApplName
        cmd.Parameters(7).Value = AppMgr
        cmd.Parameters(8).Value = Priority
        cmd.Parameters(9).Value = ReqTitle
        cmd.Parameters(10).Value = ReqDesc
        cmd.Parameters(11).Value = ExpCloseDate
        cmd.Parameters(12).Value = TeamLead
        cmd.Parameters(13).Value = Developer
        cmd.Parameters(14).Value = ExpQC
        cmd.Parameters(15).Value = EstManHour
        cmd.Parameters(16).Value = EstTotalCost
        cmd.Parameters(17).Value = ExpStartDate
        cmd.Parameters(18).Value = ExpEndDate
        cmd.Parameters(19).Value = RemarksIT
        cmd.Parameters(20).Value = ActStartDate
        cmd.Parameters(21).Value = RemarksDev
        cmd.Parameters(22).Value = ActManHour
        cmd.Parameters(23).Value = ActTotalCost
        cmd.Parameters(24).Value = ActEndDate
        cmd.Parameters(25).Value = ActQC
        cmd.Parameters(26).Value = StatusQC
        cmd.Parameters(27).Value = RemarksQC
        cmd.Parameters(28).Value = UATReadyDate
        cmd.Parameters(29).Value = UserUAT
        cmd.Parameters(30).Value = StatusUAT
        cmd.Parameters(31).Value = RemarksUAT
        cmd.Parameters(32).Value = ExpCutinDate
        cmd.Parameters(33).Value = DeployUser
        cmd.Parameters(34).Value = ActCutinDate
        cmd.Parameters(35).Value = RemarksDeploy
        cmd.Parameters(36).Value = IsUATFailed
        cmd.Parameters(37).Value = ModUser
        cmd.Parameters(38).Value = EstHourRate
        cmd.Parameters(39).Value = ActHourRate
        cmd.Parameters(40).Value = ReqNo
        cmd.Execute
        If StatusUAT = "F" And SCode = "UAT" Then
            cmd.CommandText = " INSERT INTO RequestUATDetails " & _
                                "( ReqNo, Developer, ActEndDate, ActQC, UATReadyDate, UserUAT, UATFailedDate, RemarksUAT) " & _
                                "(SELECT ReqNo, Developer, ActEndDate, ActQC, UATReadyDate, UserUAT, getdate(), RemarksUAT " & _
                                "FROM RequestDetails WHERE ReqNo=" & ReqNo & ")"
            cmd.Execute
        End If

        Dim intarr As Integer
        Dim i      As Integer
    
        intarr = UBound(marrFileName)
        
        cmd.CommandText = " DELETE FROM RequestAttach " & _
                          "WHERE ReqNo=? "
        cmd.Parameters(0).Value = ReqNo
        cmd.Execute

        If intarr > 0 Then
            cmd.CommandText = " INSERT INTO RequestAttach " & _
                            "(ReqNo, DocRefNo, DocFileName, CrUser,CrDate) " & _
                            "VALUES (?,?,?,?,?) "
            For i = 1 To intarr
                cmd.Parameters(0).Value = ReqNo
                cmd.Parameters(1).Value = marrFileName(i)
                cmd.Parameters(2).Value = marrFileNameDisplay(i)
                cmd.Parameters(3).Value = marrFileUser(i)
                cmd.Parameters(4).Value = marrFileDate(i)
                cmd.Execute
            Next
        End If
        'update status
        If intCS < 2 Then
            If Trim(TeamLead) <> "" Then
                cmd.CommandText = "UPDATE RequestDetails " & _
                                  "SET SCode='Scoping' WHERE ReqNo =?"
                cmd.Parameters(0).Value = ReqNo
                cmd.Execute
            End If
        End If
        If intCS < 3 Then
            If Trim(Me.Developer) <> "" Then
                cmd.CommandText = "UPDATE RequestDetails " & _
                                  "SET SCode='Queued' WHERE ReqNo =?"
                cmd.Parameters(0).Value = ReqNo
                cmd.Execute
            End If
        End If
        If intCS < 4 Then
            If Me.ActStartDate <> "12:00:00 AM" Then
                cmd.CommandText = "UPDATE RequestDetails " & _
                                  "SET SCode='DIP' WHERE ReqNo =?"
                cmd.Parameters(0).Value = ReqNo
                cmd.Execute
            End If
        End If
        If intCS < 5 Then
            If Me.UATReadyDate <> "12:00:00 AM" Then
                If Trim(Me.StatusQC) = "P" Then
                    If IsUATFailed = "Y" Then
                        IsUATFailed = "N"
                        cmd.CommandText = "UPDATE RequestDetails " & _
                                          "SET SCode='UAT', IsUATFailed ='N' " & _
                                          "WHERE ReqNo =?"
                    Else
                        cmd.CommandText = "UPDATE RequestDetails " & _
                                          "SET SCode='UAT' WHERE ReqNo =?"
                    End If
                    cmd.Parameters(0).Value = ReqNo
                    cmd.Execute
                End If
            End If
        End If
        If intCS < 6 Then
            If Trim(Me.UserUAT) <> "" Then
                If Trim(Me.StatusUAT) = "P" Then
                    cmd.CommandText = "UPDATE RequestDetails " & _
                                      "SET SCode='Deploy' WHERE ReqNo =?"
                    cmd.Parameters(0).Value = ReqNo
                    cmd.Execute
                Else
                    cmd.CommandText = "UPDATE RequestDetails " & _
                                      "SET SCode='DIP', StatusQC='' WHERE ReqNo =?"
                    cmd.Parameters(0).Value = ReqNo
                    cmd.Execute
                End If
            End If
        End If
        If intCS < 7 Then
            If Trim(Me.DeployUser) <> "" Then
                cmd.CommandText = "UPDATE RequestDetails " & _
                                  "SET SCode='Closed' WHERE ReqNo =?"
                cmd.Parameters(0).Value = ReqNo
                cmd.Execute
            End If
        End If

        'insert user if not exists
        cmd.CommandText = " SELECT UName FROM ASRUsers WHERE UName=? "
        cmd.Parameters(0).Value = ReqUser
        
        rst.Open cmd
        
        If rst.EOF = True Then
            cmd.CommandText = "INSERT INTO ASRUsers([UName], [Email], [Dept], [MgrName], [MgrEmail], CrUser) " & _
                              "VALUES (?,?,?,?,?,?) "
            cmd.Parameters(0).Value = ReqUser
            cmd.Parameters(1).Value = ReqEmail
            cmd.Parameters(2).Value = DeptName
            cmd.Parameters(3).Value = ReqUserMgr
            cmd.Parameters(4).Value = ReqMgrEmail
            cmd.Parameters(5).Value = ModUser
        Else
            cmd.CommandText = "UPDATE ASRUsers SET Email=?, Dept=?, MgrName=?, MgrEmail=?, ModUser=? " & _
                              " WHERE UName=? "
            cmd.Parameters(0).Value = ReqEmail
            cmd.Parameters(1).Value = DeptName
            cmd.Parameters(2).Value = ReqUserMgr
            cmd.Parameters(3).Value = ReqMgrEmail
            cmd.Parameters(4).Value = ModUser
            cmd.Parameters(5).Value = ReqUser
        End If
        cmd.Execute
        
    cnn.CommitTrans
    rst.Close

On Error GoTo ErrMail_Routine

        If intCS < 2 Then
            If Trim(TeamLead) <> "" Then
            'Scoping                 --> notify user
                strTo = ReqEmail & ", "
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                cmd.Parameters(0).Value = TeamLead
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
                cmd.Parameters(0).Value = ExpQC
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName = (SELECT AppData FROM AppValues WHERE AppCode='DevMgr')"
                rst.Open cmd
                If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
                rst.Close
                                
                strBody = "REQUEST NO : " & ReqNo & vbCrLf
                strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                strBody = strBody & "STATUS : Scoping" & vbCrLf
                strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Scoping "
                strTemp = SendMail(strTo, strCC, strSubject, strBody)
                If strTemp = "1" Then
                    
                Else
                    GoTo ErrMail_Routine
                End If
                
            End If
        End If
        If Developer <> "" Then
            strdev = Developer & ","
            strdev = Mid(strdev, 1, InStr(1, strdev, ",") - 1)
        End If
        
        If (strTeamLeadOrg <> "" And strTeamLeadOrg <> Trim(TeamLead)) Or _
           (strExpQCOrg <> "" And strExpQCOrg <> Trim(ExpQC)) Or _
           (strDeveloperOrg <> "" And strDeveloperOrg <> strdev) Then
                
            strCC = ReqEmail & ", "
            cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
            
            If (strTeamLeadOrg <> "" And strTeamLeadOrg <> Trim(TeamLead)) Then
                cmd.Parameters(0).Value = strTeamLeadOrg
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
            End If
            
            If (strExpQCOrg <> "" And strExpQCOrg <> Trim(ExpQC)) Then
                cmd.Parameters(0).Value = strExpQCOrg
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
            End If
            
            If (strDeveloperOrg <> "" And strDeveloperOrg <> strdev) Then
                cmd.Parameters(0).Value = strDeveloperOrg
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
            End If
            cmd.Parameters(0).Value = TeamLead
            rst.Open cmd
            If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
            rst.Close
            
            If strdev <> "" Then
                cmd.Parameters(0).Value = strdev
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
            End If
            
            If ExpQC <> "" Then
                cmd.Parameters(0).Value = ExpQC
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
            End If
                
            strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
            strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
            strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf & vbCrLf & vbCrLf
            strBody = strBody & "IT details has been changed. Please review the IT details. " & vbCrLf
            strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - IT Details Changed "
'WriteLog vbCrLf & "**TO**" & strTo & vbCrLf & "**SUBJECT**" & strSubject & vbCrLf & "**CONTENT**" & strBody & vbCrLf
            strTemp = SendMail(strTo, strCC, strSubject, strBody)
            If strTemp = "1" Then
                
            Else
                GoTo ErrMail_Routine
            End If
        End If
        If intCS < 3 Then
            If Trim(Me.Developer) <> "" Then
                'Queued --> notify user
                strTo = ReqEmail & ", " & ReqMgrEmail & ", "
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
            
                strdev = Developer & ","
                cmd.Parameters(0).Value = Mid(strdev, 1, InStr(1, strdev, ",") - 1)
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = TeamLead
                rst.Open cmd
                If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
                rst.Close
                strBody = "REQUEST NO : " & ReqNo & vbCrLf
                strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                strBody = strBody & "START DATE : " & ExpStartDate & vbCrLf
                strBody = strBody & "END DATE : " & ExpEndDate & vbCrLf
                strBody = strBody & "EST. MANHOUR : " & EstManHour & vbCrLf
                strBody = strBody & "COST : " & EstTotalCost & vbCrLf
                strBody = strBody & "STATUS : Queued" & vbCrLf
                strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Queued"
                strTemp = SendMail(strTo, strCC, strSubject, strBody)
                If strTemp = "1" Then
                    
                Else
                    GoTo ErrMail_Routine
                End If
                
            End If
        End If

        If intCS < 4 Then
            If Me.ActStartDate <> "12:00:00 AM" Then
                'DIP --> notify user
                strTo = ReqEmail & ", "
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                cmd.Parameters(0).Value = TeamLead
                rst.Open cmd
                If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = AppMgr
                rst.Open cmd
                If rst.EOF = False Then strCC = strCC & rst.Fields(0).Value & ", "
                rst.Close
                
                strBody = "REQUEST NO : " & ReqNo & vbCrLf
                strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                strBody = strBody & "ACTUAL START DATE: " & ActStartDate & vbCrLf & vbCrLf
                strBody = strBody & "STATUS : DIP " & vbCrLf
                strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Development Started"
                strTemp = SendMail(strTo, strCC, strSubject, strBody)
                If strTemp = "1" Then
                    
                Else
                    GoTo ErrMail_Routine
                End If
            End If
            
        End If
        If intCS < 5 Then
            If Format(Me.ActEndDate, "dd/MM/yyyy") <> Format(dteActEndDateOrg, "dd/MM/yyyy") Then
                If Me.ActEndDate <> "12:00:00 AM" Then
                    'DIP complete
                    cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                    cmd.Parameters(0).Value = TeamLead
                    rst.Open cmd
                    If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
                    rst.Close
                    
                    cmd.Parameters(0).Value = Me.ExpQC
                    rst.Open cmd
                    If rst.EOF = False Then strTo = rst.Fields(0).Value & ", "
                    rst.Close
                    
                    If ActQC <> "" Then
                        cmd.Parameters(0).Value = Me.ActQC
                        rst.Open cmd
                        If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                        rst.Close
                    End If
                    strBody = "REQUEST NO : " & ReqNo & vbCrLf
                    strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                    strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                    strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                    strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                    strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                    strBody = strBody & "ACTUAL COMPLETE DATE: " & ActEndDate & vbCrLf & vbCrLf
                    strBody = strBody & "STATUS : DIP " & vbCrLf
                    strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Development Completed"
                    strTemp = SendMail(strTo, strCC, strSubject, strBody)
                    If strTemp = "1" Then
                        
                    Else
                        GoTo ErrMail_Routine
                    End If
                End If
            End If
            If Me.UATReadyDate <> "12:00:00 AM" And Me.IsUATFailed <> "Y" Then
                'UAT' --> notify user
                strTo = ReqEmail & ", "
                
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                cmd.Parameters(0).Value = AppMgr
                rst.Open cmd
                If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = TeamLead
                rst.Open cmd
                If rst.EOF = False Then strCC = strCC & rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = ActQC
                rst.Open cmd
                If rst.EOF = False Then strCC = strCC & rst.Fields(0).Value & ", "
                rst.Close
                
                strBody = "REQUEST NO : " & ReqNo & vbCrLf
                strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                strBody = strBody & "UAT READY DATE : " & UATReadyDate & vbCrLf & vbCrLf
                strBody = strBody & "STATUS : UAT " & vbCrLf
                strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Ready for UAT"
                strTemp = SendMail(strTo, strCC, strSubject, strBody)
                If strTemp = "1" Then
                    
                Else
                    GoTo ErrMail_Routine
                End If
            End If
        End If
        If intCS < 6 Then
            strTo = ""
            If Trim(Me.UserUAT) <> "" And Me.IsUATFailed <> "Y" Then
                If Trim(Me.StatusUAT) = "P" Then
                    'Deploy'
                    cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                    cmd.Parameters(0).Value = AppMgr
                    rst.Open cmd
                    If rst.EOF = False Then strTo = rst.Fields(0).Value & ", "
                    rst.Close
                    
                    cmd.Parameters(0).Value = TeamLead
                    rst.Open cmd
                    If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                    rst.Close
                    
                    strdev = Developer & ","
                    cmd.Parameters(0).Value = Mid(strdev, 1, InStr(1, strdev, ",") - 1)
                    rst.Open cmd
                    If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                    rst.Close
        
                    strBody = "REQUEST NO : " & ReqNo & vbCrLf
                    strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                    strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                    strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                    strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                    strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                    strBody = strBody & "EXP. CUT-IN DATE: " & Format(ExpCutinDate, "dd/MM/yyyy") & vbCrLf & vbCrLf
                    strBody = strBody & "STATUS : Deploy "
                    strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Deploy"
                End If
            End If
            
            If Trim(Me.UserUAT) <> "" And Me.IsUATFailed = "Y" And Trim(Me.StatusUAT) = "F" And Trim(SCode) = "UAT" Then
                'dip'
                
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                cmd.Parameters(0).Value = AppMgr
                rst.Open cmd
                If rst.EOF = False Then strTo = rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = TeamLead
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
                
                strdev = Developer & ","
                cmd.Parameters(0).Value = Mid(strdev, 1, InStr(1, strdev, ",") - 1)
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = ActQC
                rst.Open cmd
                If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
                rst.Close
                
                strCC = ReqEmail & ", "
    
                strBody = "REQUEST NO : " & ReqNo & vbCrLf
                strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                strBody = strBody & "REMARKS : " & RemarksUAT & vbCrLf & vbCrLf
                strBody = strBody & "STATUS : DIP "
                strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - UAT Failed"
            End If
            If strTo <> "" Then
                strTemp = SendMail(strTo, strCC, strSubject, strBody)
                If strTemp = "1" Then
                    
                Else
                    GoTo ErrMail_Routine
                End If
            End If
            
        End If
        If intCS < 7 Then
            If Trim(Me.DeployUser) <> "" Then
                'Closed'                --> notify user
                strTo = ReqEmail & ", "
                
                cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
                cmd.Parameters(0).Value = AppMgr
                rst.Open cmd
                If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
                rst.Close
                
                cmd.Parameters(0).Value = TeamLead
                rst.Open cmd
                If rst.EOF = False Then strCC = strCC & rst.Fields(0).Value & ", "
                rst.Close
    
                strBody = "REQUEST NO : " & ReqNo & vbCrLf
                strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
                strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
                strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
                strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
                strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
                strBody = strBody & "ACTUAL CUT-IN DATE : " & ActCutinDate & vbCrLf & vbCrLf
                strBody = strBody & "STATUS : Closed" & vbCrLf
                strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Closed"
                strTemp = SendMail(strTo, strCC, strSubject, strBody)
                If strTemp = "1" Then
                    
                Else
                    GoTo ErrMail_Routine
                End If
                
            End If
        End If

    UpdateRecord = "OK"
    Set cmd = Nothing
    cnn.Close
    Set cnn = Nothing
    GoTo Exit_Routine
    
ErrMail_Routine:
     UpdateRecord = "ErrMail"
     Set cmd = Nothing
     cnn.Close
     Set cnn = Nothing
     Err.Raise Err.Number, Err.Source, Err.Description
     GoTo Exit_Routine
    
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
    If Err.Number <> 0 Then
        UpdateRecord = ""
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Function UpdateStatusCH(reqObj As Request) As String
'update cancel/hold
On Error GoTo Err_Routine
    
    Dim cnn      As ADODB.Connection
    Dim cmd      As ADODB.Command
    Dim intTrans As Integer
    Dim strStatus As String
    Dim strRemarks As String
    Dim strUser As String
    Dim intReqNo As Integer
    Dim rst      As ADODB.Recordset
    Dim strTo          As String
    Dim strCC          As String
    Dim strBody        As String
    Dim strSubject     As String
    Dim strTemp        As String
    
    Set cnn = SQLConnWrite
    cnn.Open
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdText
    
    strStatus = reqObj.Form("StatusCH")
    strRemarks = reqObj.Form("RemarksCH")
    strUser = reqObj.Form("UID")
    intReqNo = reqObj.Form("ReqNo")
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    
    intTrans = cnn.BeginTrans
        cmd.CommandText = "UPDATE RequestDetails " & _
                          "SET SCode=?, RemarksCH=?, ModUser=?, ModDate=getdate() " & _
                          "WHERE ReqNo =?"
        cmd.Parameters(0).Value = strStatus
        cmd.Parameters(1).Value = strRemarks
        cmd.Parameters(2).Value = strUser
        cmd.Parameters(3).Value = intReqNo
        cmd.Execute
        
    cnn.CommitTrans
   
On Error GoTo ErrMail_Routine

    strTo = ReplaceNull(Trim(reqObj.Form("reqemail")), "") & ", "
    
    cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
    cmd.Parameters(0).Value = ReplaceNull(Trim(reqObj.Form("TeamLead")), "")
    rst.Open cmd
    If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
    rst.Close
        
    cmd.Parameters(0).Value = ReplaceNull(Trim(reqObj.Form("appmgr")), "")
    rst.Open cmd
    If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
    rst.Close
    
    cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName = (SELECT AppData FROM AppValues WHERE AppCode='DevMgr')"
    rst.Open cmd
    If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
    rst.Close
    
    strBody = "REQUEST NO : " & intReqNo & vbCrLf
    strBody = strBody & "REQUESTOR : " & Trim(reqObj.Form("requestor")) & vbCrLf
    strBody = strBody & "TITLE : " & Trim(reqObj.Form("reqtitle")) & vbCrLf
    strBody = strBody & "APPLICATION : " & Trim(reqObj.Form("appl")) & vbCrLf
    strBody = strBody & "DEPT : " & Trim(reqObj.Form("dept")) & vbCrLf
    strBody = strBody & "CUSTOMER : " & Trim(reqObj.Form("customer")) & vbCrLf & vbCrLf
    strBody = strBody & "DESCRIPTION : " & Trim(reqObj.Form("Desc")) & vbCrLf & vbCrLf
    strBody = strBody & "Reason : " & Trim(reqObj.Form("RemarksCH")) & vbCrLf & vbCrLf
    strBody = strBody & "PREVIOUS STATUS : " & Trim(reqObj.Form("SCode")) & vbCrLf
    
    If UCase(strStatus) = "HOLD" Then
        strSubject = "ASR Request - ReqNo: " & intReqNo & " - " & Trim(reqObj.Form("appl")) & " - Hold"
        strBody = strBody & "CURRENT STATUS : Hold" & vbCrLf
        strBody = strBody & "Your Request is on Hold." & vbCrLf
    Else
        strSubject = "ASR Request - ReqNo: " & intReqNo & " - " & Trim(reqObj.Form("appl")) & " - Cancelled"
        strBody = strBody & "CURRENT STATUS : Cancelled" & vbCrLf
        strBody = strBody & "Your Request is cancelled." & vbCrLf
    End If

    strTemp = SendMail(strTo, strCC, strSubject, strBody)
    If strTemp = "1" Then
        
    Else
        GoTo ErrMail_Routine
    End If

    UpdateStatusCH = "OK"
    Set cmd = Nothing
    cnn.Close
    Set cnn = Nothing
    GoTo Exit_Routine
    
ErrMail_Routine:
     UpdateStatusCH = "ErrMail"
     Set cmd = Nothing
     cnn.Close
     Set cnn = Nothing
     Err.Raise Err.Number, Err.Source, Err.Description
     GoTo Exit_Routine
       
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
    If Err.Number <> 0 Then
        UpdateStatusCH = ""
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Function UpdateRelease() As String

On Error GoTo Err_Routine
    
    Dim cnn      As ADODB.Connection
    Dim cmd      As ADODB.Command
    Dim intTrans As Integer
    Dim strStatus As String
    Dim strSQL As String
    Dim rst As ADODB.Recordset
    Dim strTo          As String
    Dim strCC          As String
    Dim strBody        As String
    Dim strSubject     As String
    Dim strTemp        As String
      
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT top 1 SCode from dbo.RequestDetailsStatus " & _
                      "WHERE ReqNo=? AND SCode<>'Hold' ORDER BY CrDate DESC"
    cmd.Parameters(0).Value = ReqNo
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    If rst.EOF = True Then
        UpdateRelease = "StatusNE"
        rst.Close
    Else
        strStatus = rst.Fields(0).Value
        rst.Close
        Set rst = Nothing
        Set cmd = Nothing
        
        Set cnn = SQLConnWrite
        cnn.Open
        
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = cnn
        cmd.CommandType = adCmdText
    
        intTrans = cnn.BeginTrans
            cmd.CommandText = "UPDATE RequestDetails " & _
                              "SET SCode=?, ModUser=?, ModDate=getdate() " & _
                              "WHERE ReqNo=? "
            cmd.Parameters(0).Value = strStatus
            cmd.Parameters(1).Value = ModUser
            cmd.Parameters(2).Value = ReqNo
            cmd.Execute
        cnn.CommitTrans

On Error GoTo ErrMail_Routine
Dim strTeamLead As String
Dim strAppMgr   As String

        cmd.CommandText = "SELECT UserEmail, TeamLead, AppMgr FROM RequestDetails WHERE ReqNo=? "
        cmd.Parameters(0).Value = ReqNo
        rst.Open cmd
        If rst.EOF = False Then
            strTo = strTo & rst.Fields(0).Value & ", "
            strTeamLead = rst.Fields(1).Value
            strAppMgr = rst.Fields(2).Value
        End If
        rst.Close
        
        'strTo = ReqEmail & ", "
        
        cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName=?"
        cmd.Parameters(0).Value = strTeamLead
        rst.Open cmd
        If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
        rst.Close
            
        cmd.Parameters(0).Value = strAppMgr
        rst.Open cmd
        If rst.EOF = False Then strTo = strTo & rst.Fields(0).Value & ", "
        rst.Close
        
        cmd.CommandText = "SELECT Email FROM ASRUsers WHERE UName = (SELECT AppData FROM AppValues WHERE AppCode='DevMgr')"
        rst.Open cmd
        If rst.EOF = False Then strCC = rst.Fields(0).Value & ", "
        rst.Close
        
        strBody = "REQUEST NO : " & ReqNo & vbCrLf
        strBody = strBody & "REQUESTOR : " & ReqUser & vbCrLf
        strBody = strBody & "TITLE : " & ReqTitle & vbCrLf
        strBody = strBody & "APPLICATION : " & ApplName & vbCrLf
        strBody = strBody & "DEPT : " & Me.DeptName & vbCrLf
        strBody = strBody & "CUSTOMER : " & Me.CName & vbCrLf & vbCrLf
        strBody = strBody & "DESCRIPTION : " & ReqDesc & vbCrLf & vbCrLf
        
        strSubject = "ASR Request - ReqNo: " & ReqNo & " - " & ApplName & " - Release"
        strBody = strBody & "CURRENT STATUS : " & strStatus & vbCrLf

        strTemp = SendMail(strTo, strCC, strSubject, strBody)
        If strTemp = "1" Then
            
        Else
            GoTo ErrMail_Routine
        End If

        UpdateRelease = "OK"
        Set cmd = Nothing
        cnn.Close
        Set cnn = Nothing
        GoTo Exit_Routine
    
ErrMail_Routine:
        UpdateRelease = "ErrMail"
        Set cmd = Nothing
        cnn.Close
        Set cnn = Nothing
        Err.Raise Err.Number, Err.Source, Err.Description
        GoTo Exit_Routine
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    Set cnn = Nothing
       
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    If intTrans = 1 Then
        cnn.RollbackTrans
        intTrans = intTrans - 1
    End If
    If Not cnn Is Nothing Then
        If cnn.State = adStateOpen Then cnn.Close
        Set cnn = Nothing
    End If
    If Err.Number <> 0 Then
        UpdateRelease = ""
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Function RetrieveStatusNo(ByVal strStatus As String) As Integer

On Error GoTo Err_Routine
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT SeqNo FROM RequestStatus WHERE SCode =?"
    cmd.Parameters(0).Value = strStatus
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    RetrieveStatusNo = rst.Fields(0).Value
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveRecord() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    Dim intRate As Integer
    Dim cls As clsList
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT ReqNo, SCode, ReqUser, UserEmail, ReqUserMgr, UserMgrEmail, CName, DeptName, " & _
                        "ApplName, AppMgr, Priority, ReqTitle, ReqDesc, " & _
                        "CASE CONVERT(CHAR(12),ExpCloseDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ExpCloseDate,103) END   As ExpCloseDate, " & _
                        "RemarksCH, TeamLead, Developer, ExpQC, EstManHour, EstHourRate, EstTotalCost, " & _
                        "CASE CONVERT(CHAR(12),ExpStartDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ExpStartDate,103) END   As ExpStartDate, " & _
                        "CASE CONVERT(CHAR(12),ExpEndDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ExpEndDate,103) END   As ExpEndDate, " & _
                        "RemarksIT, CASE CONVERT(CHAR(12),ActStartDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ActStartDate,103) END   As ActStartDate, RemarksDev, " & _
                        "ActManHour, ActHourRate, ActTotalCost, " & _
                        "CASE CONVERT(CHAR(12),ActEndDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ActEndDate,103) END   As ActEndDate, ActQC, StatusQC, RemarksQC, " & _
                        "CASE CONVERT(CHAR(12),UATReadyDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),UATReadyDate,103) END   As UATReadyDate, " & _
                        "UserUAT, StatusUAT, RemarksUAT, " & _
                        "CASE CONVERT(CHAR(12),ExpCutinDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ExpCutinDate,103) END   As ExpCutinDate, " & _
                        "DeployUser,  " & _
                        "CASE CONVERT(CHAR(12),ActCutinDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ActCutinDate,103) END   As ActCutinDate, " & _
                        "RemarksDeploy, IsUATFailed, NoUATFailed, CrUser, CrDate, ModUser, ModDate, ReqType " & _
                        "FROM RequestDetails WHERE ReqNo=? "
   
    cmd.Parameters(0).Value = ReqNo
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    If rst.EOF = False Then
        SCode = rst.Fields("SCode").Value
        ReqUser = rst.Fields("ReqUser").Value
        ReqEmail = rst.Fields("UserEmail").Value
        ReqUserMgr = rst.Fields("ReqUserMgr").Value
        ReqMgrEmail = rst.Fields("UserMgrEmail").Value
        CName = rst.Fields("CName").Value
        DeptName = rst.Fields("DeptName").Value
        ApplName = rst.Fields("ApplName").Value
        AppMgr = rst.Fields("AppMgr").Value
        Priority = rst.Fields("Priority").Value
        ReqTitle = rst.Fields("ReqTitle").Value
        ReqDesc = rst.Fields("ReqDesc").Value
        ExpCloseDateStr = Trim(rst.Fields("ExpCloseDate").Value)
        RemarksCH = ReplaceNull(rst.Fields("RemarksCH").Value, "")
        TeamLead = ReplaceNull(rst.Fields("TeamLead").Value, "")
        Developer = ReplaceNull(rst.Fields("Developer").Value, "")
        ExpQC = ReplaceNull(rst.Fields("ExpQC").Value, "")
        EstManHour = ReplaceNull(rst.Fields("EstManHour").Value, 0)
        intRate = ReplaceNull(rst.Fields("EstHourRate").Value, 0)
        If intRate = 0 Then
            Set cls = New clsList
            If rst.Fields("ReqType").Value = "P" Then
                intRate = cls.RetrieveHourRate("Proj")
            Else
                intRate = cls.RetrieveHourRate()
            End If
            EstHourRate = intRate
            ActHourRate = intRate
            
            Set cls = Nothing
        Else
            EstHourRate = intRate
            intRate = ReplaceNull(rst.Fields("ActHourRate").Value, 0)
            If intRate = 0 Then
                Set cls = New clsList
                If rst.Fields("ReqType").Value = "P" Then
                    intRate = cls.RetrieveHourRate("Proj")
                Else
                    intRate = cls.RetrieveHourRate()
                End If
                Set cls = Nothing
            End If
            ActHourRate = intRate
        End If
        EstTotalCost = ReplaceNull(rst.Fields("EstTotalCost").Value, 0)
        ExpStartDateStr = Trim(ReplaceNull(rst.Fields("ExpStartDate").Value, ""))
        ExpEndDateStr = Trim(ReplaceNull(rst.Fields("ExpEndDate").Value, ""))
        RemarksIT = ReplaceNull(rst.Fields("RemarksIT").Value, "")
        ActStartDateStr = Trim(ReplaceNull(rst.Fields("ActStartDate").Value, ""))
        RemarksDev = ReplaceNull(rst.Fields("RemarksDev").Value, "")
        ActManHour = ReplaceNull(rst.Fields("ActManHour").Value, 0)
        ActTotalCost = ReplaceNull(rst.Fields("ActTotalCost").Value, 0)
        ActEndDateStr = Trim(ReplaceNull(rst.Fields("ActEndDate").Value, ""))
        ActQC = ReplaceNull(rst.Fields("ActQC").Value, "")
        StatusQC = ReplaceNull(rst.Fields("StatusQC").Value, "")
        RemarksQC = ReplaceNull(rst.Fields("RemarksQC").Value, "")
        UATReadyDateStr = Trim(ReplaceNull(rst.Fields("UATReadyDate"), ""))
        UserUAT = ReplaceNull(rst.Fields("UserUAT").Value, "")
        StatusUAT = ReplaceNull(rst.Fields("StatusUAT").Value, "")
        RemarksUAT = ReplaceNull(rst.Fields("RemarksUAT").Value, "")
        ExpCutinDateStr = Trim(ReplaceNull(rst.Fields("ExpCutinDate").Value, ""))
        DeployUser = ReplaceNull(rst.Fields("DeployUser").Value, "")
        ActCutinDateStr = Trim(ReplaceNull(rst.Fields("ActCutinDate").Value, ""))
        RemarksDeploy = ReplaceNull(rst.Fields("RemarksDeploy").Value, "")
        IsUATFailed = ReplaceNull(rst.Fields("IsUATFailed").Value, "")
        NoUATFailed = ReplaceNull(rst.Fields("NoUATFailed").Value, 0)
        CrUser = ReplaceNull(rst.Fields("CrUser").Value, "")
        CrDate = ReplaceNull(rst.Fields("CrDate").Value, "0")
        ModUser = ReplaceNull(rst.Fields("ModUser").Value, "")
        ModDate = ReplaceNull(rst.Fields("ModDate").Value, "0")
        ReqType = ReplaceNull(rst.Fields("ReqType").Value, "")
    End If
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveAttachements() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT DocRefNo, ReqNo, DocFileName, CrUser, CrDate, ModDate " & _
                        "FROM RequestAttach WHERE ReqNo = ? "
   
    cmd.Parameters(0).Value = ReqNo
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveAttachements = rst
    
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

    ReqUser = ReplaceNull(Trim(reqObj.Form("requestor")), "")
    ReqEmail = ReplaceNull(Trim(reqObj.Form("reqemail")), "")
    ReqUserMgr = ReplaceNull(Trim(reqObj.Form("reqmgr")), "")
    ReqMgrEmail = ReplaceNull(Trim(reqObj.Form("reqmgremail")), "")
    CName = ReplaceNull(Trim(reqObj.Form("customer")), "")
    DeptName = ReplaceNull(Trim(reqObj.Form("dept")), "")
    ApplName = ReplaceNull(Trim(reqObj.Form("appl")), "")
    AppMgr = ReplaceNull(Trim(reqObj.Form("appmgr")), "")
    Priority = ReplaceNull(Trim(reqObj.Form("priority")), 3)
    ReqTitle = ReplaceNull(Trim(reqObj.Form("reqtitle")), "")
    ReqDesc = ReplaceNull(Trim(reqObj.Form("Desc")), "")
    'ExpCloseDate = Format(ReplaceNull(Trim(reqObj.Form("expclosedate")), "0"), "dd/MM/yyyy")
    ExpCloseDateStr = ReplaceNull(Trim(reqObj.Form("expclosedate")), "0")
    If ExpCloseDateStr = "" Then
        ExpCloseDate = "0"
    Else
        ExpCloseDate = Format(ExpCloseDateStr, "dd/MM/yyyy")
    End If
        
    ReqType = ReplaceNull(Trim(reqObj.Form("reqtype")), "")
    CrUser = Trim(reqObj.Form("UID"))
    
    Dim intarr As Integer
    Dim i      As Integer
    
    intarr = reqObj.Form("fname").Count - 1
    ReDim marrFileName(intarr)
    ReDim marrFileNameDisplay(intarr)
    ReDim marrFileUser(intarr)
    ReDim marrFileDate(intarr)
    For i = 2 To intarr + 1
        marrFileName(i - 1) = reqObj.Form("fname")(i)
        marrFileNameDisplay(i - 1) = reqObj.Form("fnamedis")(i)
        marrFileUser(i - 1) = reqObj.Form("fuser")(i)
        marrFileDate(i - 1) = reqObj.Form("fdate")(i)
    Next
    
    'For i = 1 To intarr
    '    marrFileName(i) = reqObj.Form("fname")(i)
    '    marrFileNameDisplay(i) = reqObj.Form("fnamedis")(i)
    '    marrFileUser(i) = reqObj.Form("fuser")(i)
    '    marrFileDate(i) = reqObj.Form("fdate")(i)
    'Next
    
Exit_Routine:
    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub SetValuesUpd(reqObj As Request)
On Error GoTo Err_Routine

    ReqNo = Trim(reqObj.Form("ReqNo"))
    SCode = Trim(reqObj.Form("SCode"))
    ReqUser = ReplaceNull(Trim(reqObj.Form("requestor")), "")
    ReqEmail = ReplaceNull(Trim(reqObj.Form("reqemail")), "")
    ReqUserMgr = ReplaceNull(Trim(reqObj.Form("reqmgr")), "")
    ReqMgrEmail = ReplaceNull(Trim(reqObj.Form("reqmgremail")), "")
    CName = ReplaceNull(Trim(reqObj.Form("customer")), "")
    DeptName = ReplaceNull(Trim(reqObj.Form("dept")), "")
    ApplName = ReplaceNull(Trim(reqObj.Form("appl")), "")
    AppMgr = ReplaceNull(Trim(reqObj.Form("appmgr")), "")
    Priority = ReplaceNull(Trim(reqObj.Form("priority")), 3)
    ReqTitle = ReplaceNull(Trim(reqObj.Form("reqtitle")), "")
    ReqDesc = ReplaceNull(Trim(reqObj.Form("Desc")), "")
    
    ExpCloseDateStr = ReplaceNull(Trim(reqObj.Form("expclosedate")), "0")
    If ExpCloseDateStr = "" Then
        ExpCloseDate = "0"
    Else
        ExpCloseDate = Format(ExpCloseDateStr, "dd/MM/yyyy")
    End If
    TeamLead = ReplaceNull(Trim(reqObj.Form("TeamLead")), "")
    Developer = ReplaceNull(Trim(reqObj.Form("Developer")), "")
    ExpQC = ReplaceNull(Trim(reqObj.Form("ExpQC")), "")
    EstManHour = ReplaceNull(Trim(reqObj.Form("EstManHour")), 0)
    EstHourRate = ReplaceNull(Trim(reqObj.Form("EstHourRate")), 0)
    EstTotalCost = ReplaceNull(Trim(reqObj.Form("EstTotalCost")), 0)
    'ExpStartDate = ReplaceNull(Trim(reqObj.Form("ExpStartDate")), "0")
    ExpStartDateStr = ReplaceNull(Trim(reqObj.Form("ExpStartDate")), "0")
    If ExpStartDateStr = "" Then
        ExpStartDate = "0"
    Else
        ExpStartDate = Format(ExpStartDateStr, "dd/MM/yyyy")
    End If
    'ExpEndDate = ReplaceNull(Trim(reqObj.Form("ExpEndDate")), "0")
    ExpEndDateStr = ReplaceNull(Trim(reqObj.Form("ExpEndDate")), "0")
    If ExpEndDateStr = "" Then
        ExpEndDate = "0"
    Else
        ExpEndDate = Format(ExpEndDateStr, "dd/MM/yyyy")
    End If
    RemarksIT = ReplaceNull(Trim(reqObj.Form("RemarksIT")), "")
    'ActStartDate = ReplaceNull(Trim(reqObj.Form("ActStartDate")), "0")
    ActStartDateStr = ReplaceNull(Trim(reqObj.Form("ActStartDate")), "0")
    If ActStartDateStr = "" Then
        ActStartDate = "0"
    Else
        ActStartDate = Format(ActStartDateStr, "dd/MM/yyyy")
    End If
    RemarksDev = ReplaceNull(Trim(reqObj.Form("RemarksDev")), "")

    ActManHour = ReplaceNull(Trim(reqObj.Form("ActManHour")), 0)
    ActHourRate = ReplaceNull(Trim(reqObj.Form("ActHourRate")), 0)
    ActTotalCost = ReplaceNull(Trim(reqObj.Form("ActTotalCost")), 0)
    'ActEndDate = ReplaceNull(Trim(reqObj.Form("ActEndDate")), "0")
    ActEndDateStr = ReplaceNull(Trim(reqObj.Form("ActEndDate")), "0")
    If ActEndDateStr = "" Then
        ActEndDate = "0"
    Else
        ActEndDate = Format(ActEndDateStr, "dd/MM/yyyy")
    End If
    ActQC = ReplaceNull(Trim(reqObj.Form("ActQC")), "")
    StatusQC = ReplaceNull(Trim(reqObj.Form("StatusQC")), "")
    RemarksQC = ReplaceNull(Trim(reqObj.Form("RemarksQC")), "")
    'UATReadyDate = ReplaceNull(Trim(reqObj.Form("UATReadyDate")), "0")
    UATReadyDateStr = ReplaceNull(Trim(reqObj.Form("UATReadyDate")), "0")
    If UATReadyDateStr = "" Then
        UATReadyDate = "0"
    Else
        UATReadyDate = Format(UATReadyDateStr, "dd/MM/yyyy")
    End If
    UserUAT = ReplaceNull(Trim(reqObj.Form("UserUAT")), "")
    StatusUAT = ReplaceNull(Trim(reqObj.Form("StatusUAT")), "")
    RemarksUAT = ReplaceNull(Trim(reqObj.Form("RemarksUAT")), "")
    'ExpCutinDate = ReplaceNull(Trim(reqObj.Form("ExpCutinDate")), "0")
    ExpCutinDateStr = ReplaceNull(Trim(reqObj.Form("ExpCutinDate")), "")
    If ExpCutinDateStr = "" Then
        ExpCutinDate = "0"
    Else
        ExpCutinDate = Format(ExpCutinDateStr, "dd/MM/yyyy")
    End If
    IsUATFailed = ReplaceNull(Trim(reqObj.Form("UATFailed")), "")
    DeployUser = ReplaceNull(Trim(reqObj.Form("DeployUser")), "")
    'ActCutinDate = ReplaceNull(Trim(reqObj.Form("ActCutinDate")), "0")
    ActCutinDateStr = ReplaceNull(Trim(reqObj.Form("ActCutinDate")), "0")
    If ActCutinDateStr = "" Then
        ActCutinDate = "0"
    Else
        ActCutinDate = Format(ActCutinDateStr, "dd/MM/yyyy")
    End If
    RemarksDeploy = ReplaceNull(Trim(reqObj.Form("RemarksDeploy")), "")
    ModUser = Trim(reqObj.Form("UID"))

    Dim intarr As Integer
    Dim i      As Integer
    
    intarr = reqObj.Form("fname").Count - 1
    ReDim marrFileName(intarr)
    ReDim marrFileNameDisplay(intarr)
    ReDim marrFileUser(intarr)
    ReDim marrFileDate(intarr)
    For i = 2 To intarr + 1
        marrFileName(i - 1) = reqObj.Form("fname")(i)
        marrFileNameDisplay(i - 1) = reqObj.Form("fnamedis")(i)
        marrFileUser(i - 1) = reqObj.Form("fuser")(i)
        marrFileDate(i - 1) = reqObj.Form("fdate")(i)
    Next
        
Exit_Routine:
    Exit Sub
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
