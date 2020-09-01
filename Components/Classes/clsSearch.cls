VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RetrieveSearch(ByVal strCName As String, _
                               ByVal strApplName As String, _
                               ByVal strStatus As String, _
                               ByVal strOrder As String, _
                               ByVal strSort As String, _
                               ByVal lngReqNo As Long, _
                               ByVal strDeveloper As String, _
                               ByVal strAppMgr As String, _
                               ByVal strTeamLead As String, _
                               ByVal strReqUser As String, _
                               ByVal strDateFrom As String, _
                               ByVal strDateTo As String, _
                               ByVal strReqType As String) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    strSQL = "SELECT  Convert(char(4),Replicate('0', 4 - Len(ReqNo)) + Convert(varchar(4),ReqNo)) As ReqNo , " & _
             "CName As Customer, DeptName As Dept, ApplName As Application, " & _
             "ReqUser As Requestor, ReqTitle As RequestTitle, Priority, SCode As Status," & _
             "CASE CONVERT(CHAR(12),ExpStartDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ExpStartDate,103) END   As ExpStartDate, " & _
             "CASE CONVERT(CHAR(12),ExpEndDate,103) WHEN '30/12/1899' THEN '' ELSE convert(char(12),ExpEndDate,103) END   As ExpComplDate, " & _
             "AppMgr As ITAccountMananger, " & _
             "TeamLead As TeamLeader, Developer, CrDate as ReqDate, ReqType " & _
             "FROM RequestDetails WHERE 1=1 "
    
    If strCName <> "" Then strSQL = strSQL & " AND CName LIKE '" & strCName & "%' "
    If strApplName <> "" Then strSQL = strSQL & " AND ApplName LIKE '" & strApplName & "%' "
    If lngReqNo <> 0 Then strSQL = strSQL & " AND ReqNo = " & lngReqNo
    '
    'If strStatus <> "" Then
    '    If Trim(strStatus) <> "Closed" Then
    '        strSQL = strSQL & " AND SCode  = '" & strStatus & "' "
    '    Else
    '        strSQL = strSQL & " AND SCode='Closed' AND CONVERT(Char(8),[CrDate],112) > CONVERT(Char(8),getdate()-30,112)"
    '    End If
    'Else
    '    strSQL = strSQL & " AND (SCode<>'Closed' OR (SCode='Closed' AND convert(char(8),[CrDate],112) > convert(char(8),getdate()-30,112))) "
    'End If
    '2 jul 2003
    If Trim(strStatus) <> "" Then strSQL = strSQL & " AND SCode  = '" & strStatus & "' "
    If strDeveloper <> "" Then strSQL = strSQL & " AND Developer  = '" & strDeveloper & "' "
    If strAppMgr <> "" Then strSQL = strSQL & " AND AppMgr  = '" & strAppMgr & "' "
    If strTeamLead <> "" Then strSQL = strSQL & " AND TeamLead  = '" & strTeamLead & "' "
    If strReqUser <> "" Then strSQL = strSQL & " AND ReqUser  = '" & strReqUser & "' "
    If strDateFrom <> "" Then strSQL = strSQL & " AND CONVERT(CHAR(8), CrDate , 112) >= '" & Format(strDateFrom, "yyyyMMdd") & "' "
    If strDateTo <> "" Then strSQL = strSQL & " AND CONVERT(CHAR(8), CrDate , 112) <= '" & Format(strDateTo, "yyyyMMdd") & "' "

    If strReqType <> "" Then strSQL = strSQL & " AND ReqType  = '" & strReqType & "' "
    
    If strOrder <> "" Then strSQL = strSQL & "  ORDER BY " & strOrder & " " & strSort

    cmd.CommandText = strSQL
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    'Set RetrieveSearch = cmd.Execute
    Set RetrieveSearch = rst
    'rst.Close
    
    Set rst = Nothing
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

'"SELECT CName, DeptName, ReqUser, ReqTitle, ExpStartDate, " & _
"datediff("d",ExpStartDate,getdate()), R.SCode, TeamLead, Developer " & _
"FROM dbo.RequestDetails R INNER JOIN RequestStatus S on R.SCode=S.SCode " & _
"WHERE S.SeqNo between 2 and 4 "


