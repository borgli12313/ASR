VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RetrieveStatusReport() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT S.SCode AS Status, COUNT(R.SCode) As NoReq, S.SeqNo FROM RequestDetails R " & _
                      "RIGHT OUTER JOIN RequestStatus S on S.SCode=R.SCode " & _
                      "GROUP BY S.SCode, S.SeqNo " & _
                      "ORDER BY S.SeqNo"
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    
    Set RetrieveStatusReport = rst
    
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

Public Function RetrieveReqNoStatusReport(ByVal intReqNo As Integer) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT R.SCode, R.CrUser, R.CrDate  FROM RequestDetailsStatus R " & _
                      "INNER JOIN RequestStatus S ON S.SCode = R.SCode " & _
                      "WHERE R.ReqNo = ? " & _
                      "ORDER BY R.CrDate, S.SeqNo"
    
    cmd.Parameters(0).Value = intReqNo
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveReqNoStatusReport = rst
    
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

Public Function RetrieveStatusSumReport(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT convert(char(6),(R.CrDate),112) ,datename(""m"",R.CrDate) AS MONTH, " & _
                        "SUM(CASE R.SCode WHEN 'Open' THEN 1 ELSE 0 END) AS '  OPEN  ' , " & _
                        "SUM(CASE R.SCode WHEN 'Scoping' THEN 1 ELSE 0 END) AS SCOPING , " & _
                        "SUM(CASE R.SCode WHEN 'Queued' THEN 1 ELSE 0 END) AS QUEUED , " & _
                        "SUM(CASE R.SCode WHEN 'DIP' THEN 1 ELSE 0 END) AS '  DIP  ' , " & _
                        "SUM(CASE R.SCode WHEN 'UAT' THEN 1 ELSE 0 END) AS '  UAT  ' , " & _
                        "SUM(CASE R.SCode WHEN 'Deploy' THEN 1 ELSE 0 END) AS DEPLOY , " & _
                        "SUM(CASE R.SCode WHEN 'Closed' THEN 1 ELSE 0 END) AS CLOSED , " & _
                        "SUM(CASE R.SCode WHEN 'Cancelled' THEN 1 ELSE 0 END) AS CANCELLED , " & _
                        "SUM(CASE R.SCode WHEN 'Hold' THEN 1 ELSE 0 END) AS ' HOLD '  ,     " & _
                        "count(r.scode) As TOTAL " & _
                        "From RequestDetails R  " & _
                        "WHERE convert(char(8),(R.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "GROUP BY convert(char(6),(R.CrDate),112), datename(""m"",R.CrDate) " & _
                        "ORDER BY 1 "
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveStatusSumReport = rst
    
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

Public Function RetrieveBacklogReport() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rst As ADODB.Recordset
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT CName as Customer, DeptName as Dept, ReqUser as Requestor, " & _
                      "ReqTitle As Title, Priority, R.CrDate as RequestDate,  " & _
                      "IsNull(CASE Convert(char(10), ExpStartDate, 103) WHEN '30/12/1899' THEN '' ELSE " & _
                      "Convert(char(10), ExpStartDate,103) END , '') AS ExpStartDate,  " & _
                      "IsNull( CASE Convert(char(10), ExpEndDate, 103) WHEN '30/12/1899' THEN '' ELSE " & _
                      "Convert(char(10), ExpEndDate,103) END, '') As ExpComplDate, " & _
                      "R.SCode as Status, AppMgr as AccountManager, TeamLead, Developer, ReqNo " & _
                      "FROM dbo.RequestDetails R INNER JOIN RequestStatus S on R.SCode=S.SCode " & _
                      "Where S.SeqNo < 4 " & _
                      "ORDER BY Priority, S.SeqNo, ReqNo"
                      
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveBacklogReport = rst
    
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

Public Function RetrieveQualilyKPIReport(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT CONVERT(char(6), (S.CrDate),112) ,DATENAME(""m"",S.CrDate) AS MONTH, " & _
                        "Count(R.ReqNo) As TotalUAT, " & _
                        "SUM(CASE IsNull(R.NoUATFailed,0) WHEN 0 THEN 1 ELSE 0 End) As UATPass, " & _
                        "SUM(CASE IsNull(R.NoUATFailed,0) WHEN 0 THEN 0 ELSE 1 End) As UATFail, " & _
                        "Convert(Decimal(5,2), " & _
                        "(SUM(CASE IsNull(R.NoUATFailed,0) WHEN 0 THEN 1 ELSE 0 End)* 100 )/ " & _
                        " Convert(Decimal(5,2),Count(R.ReqNo))) As '%Pass' " & _
                        "FROM dbo.RequestDetails R " & _
                        "INNER JOIN RequestDetailsStatus S ON S.ReqNo=R.ReqNo " & _
                        "WHERE S.SCode='Deploy'" & _
                        "AND Convert(Char(8),(S.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "GROUP BY convert(char(6),(S.CrDate),112), datename(""m"",S.CrDate) " & _
                        "ORDER BY 1 "
         
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveQualilyKPIReport = rst
    
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

Public Function RetrieveQualilyKPIDets(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT CONVERT(char(6), (S.CrDate),112), DATENAME(""m"",S.CrDate) AS MONTH, " & _
                        "R.ReqNo, convert(char(10),max(s2.crdate),103) As FailedDate,  " & _
                        "NoUATFailed, TeamLead, Developer, ActQC, UserUAT As UATUser " & _
                        "FROM dbo.RequestDetails R " & _
                        "INNER JOIN RequestDetailsStatus S ON S.ReqNo=R.ReqNo " & _
                        "INNER JOIN RequestDetailsStatus S2 ON S2.ReqNo=R.ReqNo " & _
                        "WHERE S.SCode='Deploy' " & _
                        "AND S2.SCode='DIP' " & _
                        "AND Convert(Char(8),(S.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "And IsNull(R.NoUATFailed,0)>0 " & _
                        "GROUP BY CONVERT(char(6), (S.CrDate),112), DATENAME(""m"",S.CrDate), " & _
                        "R.ReqNo, NoUATFailed, TeamLead, Developer, ActQC, UserUAT " & _
                        "ORDER BY 1 "
             
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveQualilyKPIDets = rst
    
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

Public Function RetrievePerKPIReport(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT CONVERT(char(6), (S.CrDate),112) ,DATENAME(""m"",S.CrDate) AS MONTH, " & _
                        "Count(R.ReqNo) As TotalUAT, " & _
                        "SUM(CASE WHEN ActEndDate>ExpEndDate THEN 0 ELSE 1 End) As MetTargetDate, " & _
                        "SUM(CASE WHEN ActEndDate>ExpEndDate THEN 1 ELSE 0 End) As ExceededTargetDate, " & _
                        "Convert(Decimal(5,2), " & _
                        "(SUM(CASE WHEN ActEndDate>ExpEndDate THEN 0 ELSE 1 End)* 100 )/" & _
                        " Convert(Decimal(5,2),Count(R.ReqNo))) As '%Pass' " & _
                        "FROM dbo.RequestDetails R " & _
                        "INNER JOIN RequestDetailsStatus S ON S.ReqNo=R.ReqNo " & _
                        "WHERE S.SCode='Deploy'" & _
                        "AND Convert(Char(8),(S.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "GROUP BY convert(char(6),(S.CrDate),112), datename(""m"",S.CrDate) " & _
                        "ORDER BY 1 "
         
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrievePerKPIReport = rst
    
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

Public Function RetrievePerKPIDets(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT CONVERT(char(6), (S.CrDate),112), DATENAME(""m"",S.CrDate) AS MONTH, " & _
                        "R.ReqNo, R.ReqTitle, ExpEndDate As ExpComplDt, ActEndDate As ActComplDt,  " & _
                        "DateDiff(""d"",ExpEndDate,ActEndDate) As DaysExceeded,  " & _
                        "TeamLead, Developer " & _
                        "FROM dbo.RequestDetails R " & _
                        "INNER JOIN RequestDetailsStatus S ON S.ReqNo=R.ReqNo " & _
                        "WHERE S.SCode='Deploy'" & _
                        "AND Convert(Char(8),(S.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "AND ActEndDate>ExpEndDate " & _
                        "ORDER BY 1 "
         
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrievePerKPIDets = rst
    
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

Public Function RetrieveCustVolStatusReport(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT CName AS Customer, " & _
                        "SUM(CASE R.SCode WHEN 'Open' THEN 1 ELSE 0 END) AS OPENREQ , " & _
                        "SUM(CASE R.SCode WHEN 'Scoping' THEN 1 ELSE 0 END) AS SCOPING , " & _
                        "SUM(CASE R.SCode WHEN 'Queued' THEN 1 ELSE 0 END) AS QUEUED , " & _
                        "SUM(CASE R.SCode WHEN 'DIP' THEN 1 ELSE 0 END) AS DIP , " & _
                        "SUM(CASE R.SCode WHEN 'UAT' THEN 1 ELSE 0 END) AS UAT , " & _
                        "SUM(CASE R.SCode WHEN 'Deploy' THEN 1 ELSE 0 END) AS DEPLOY , " & _
                        "SUM(CASE R.SCode WHEN 'Closed' THEN 1 ELSE 0 END) AS CLOSED , " & _
                        "SUM(CASE R.SCode WHEN 'Cancelled' THEN 1 ELSE 0 END) AS CANCELLED , " & _
                        "SUM(CASE R.SCode WHEN 'Hold' THEN 1 ELSE 0 END) AS HOLD ,     " & _
                        "Count(R.SCode) As TOTAL " & _
                        "From RequestDetails R  " & _
                        "WHERE Convert(Char(8),(R.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "GROUP BY CName "
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveCustVolStatusReport = rst
    
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

Public Function RetrieveBAXVolStatusReport(ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT DeptName, " & _
                        "SUM(CASE R.SCode WHEN 'Open' THEN 1 ELSE 0 END) AS OPENREQ, " & _
                        "SUM(CASE R.SCode WHEN 'Scoping' THEN 1 ELSE 0 END) AS SCOPING , " & _
                        "SUM(CASE R.SCode WHEN 'Queued' THEN 1 ELSE 0 END) AS QUEUED , " & _
                        "SUM(CASE R.SCode WHEN 'DIP' THEN 1 ELSE 0 END) AS DIP , " & _
                        "SUM(CASE R.SCode WHEN 'UAT' THEN 1 ELSE 0 END) AS UAT , " & _
                        "SUM(CASE R.SCode WHEN 'Deploy' THEN 1 ELSE 0 END) AS DEPLOY , " & _
                        "SUM(CASE R.SCode WHEN 'Closed' THEN 1 ELSE 0 END) AS CLOSED , " & _
                        "SUM(CASE R.SCode WHEN 'Cancelled' THEN 1 ELSE 0 END) AS CANCELLED , " & _
                        "SUM(CASE R.SCode WHEN 'Hold' THEN 1 ELSE 0 END) AS HOLD ,     " & _
                        "Count(R.SCode) As TOTAL " & _
                        "From RequestDetails R  " & _
                        "WHERE Convert(Char(8),(R.CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  " & _
                        "AND CName = 'BAX' " & _
                        "GROUP BY DeptName "
        
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveBAXVolStatusReport = rst
    
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

Public Function RetrieveCustDets(ByVal strCName As String, ByVal dteFrom As Date, ByVal dteTo As Date) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT ApplName As Application, ReqNo, ReqTitle As Title, ReqDesc As Description, " & _
                        "ReqUser As Requestor, DeptName, CrDate As RequestDate,  " & _
                        "CASE CONVERT(CHAR(12),ActStartDate,103) WHEN '30/12/1899' THEN '' ELSE  " & _
                        "CONVERT(char(12),ActStartDate,103) END   As ActStartDate,  " & _
                        "CASE CONVERT(CHAR(12),ActEndDate,103) WHEN '30/12/1899' THEN '' ELSE  " & _
                        "CONVERT(char(12),ActEndDate,103) END   As ActComplDt,   " & _
                        "CASE CONVERT(CHAR(12),ActCutinDate,103) WHEN '30/12/1899' THEN '' ELSE  " & _
                        "CONVERT(char(12),ActCutinDate,103) END   As ActCutinDate,  " & _
                        "IsNull(ActManHour,0) As ActManHour , SCode As Status, Priority " & _
                        "FROM RequestDetails  " & _
                        "WHERE CName=? " & _
                        "AND Convert(Char(8),(CrDate),112) BETWEEN '" & Format(dteFrom, "yyyymmdd") & _
                        "' AND '" & Format(dteTo, "yyyymmdd") & "'  "
    
    cmd.Parameters(0).Value = strCName
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    Set RetrieveCustDets = rst
    
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

Public Function RetrieveIntCost(ByVal strDate As String) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim strdt As String
    
    strdt = Year(Now()) & Format(Month(Now()) - 1, "00")
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdStoredProc
    
    If strdt = strDate Then
        cmd.CommandText = "dbo.srptIntCost"
    Else
        cmd.CommandText = "dbo.srptIntCostHis"
        cmd.Parameters(1).Value = strDate
    End If
    Set RetrieveIntCost = cmd.Execute
     
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveIntCostOthers(ByVal strDate As String) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim strdt As String
    
    strdt = Year(Now()) & Format(Month(Now()) - 1, "00")
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdStoredProc
    
    If strdt = strDate Then
        cmd.CommandText = "dbo.srptIntCostOthers"
    Else
        cmd.CommandText = "dbo.srptIntCostOthersHis"
        cmd.Parameters(1).Value = strDate
    End If
    Set RetrieveIntCostOthers = cmd.Execute
     
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function
Public Function RetrieveIntCostBAX(ByVal strDate As String) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim strdt As String
    
    strdt = Year(Now()) & Format(Month(Now()) - 1, "00")
   
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdStoredProc
    
    If strdt = strDate Then
        cmd.CommandText = "dbo.srptIntCostBAX"
    Else
        cmd.CommandText = "dbo.srptIntCostBAXHis"
        cmd.Parameters(1).Value = strDate
    End If
    Set RetrieveIntCostBAX = cmd.Execute
     
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveAgingReport() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.srptAging"
      
    Set RetrieveAgingReport = cmd.Execute
     
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveAgingReportGrp() As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.srptAgingRepGrp"
      
    Set RetrieveAgingReportGrp = cmd.Execute
     
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

