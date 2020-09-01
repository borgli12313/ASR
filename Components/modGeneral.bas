Attribute VB_Name = "modGeneral"
Option Explicit

Public Function SQLConnWrite() As ADODB.Connection
    Dim cnnWrite As ADODB.Connection
    Set cnnWrite = New ADODB.Connection
    cnnWrite.ConnectionString = CnnStrSQL
    cnnWrite.Mode = adModeReadWrite
    Set SQLConnWrite = cnnWrite
    Set cnnWrite = Nothing
End Function

Public Property Get CnnStrSQL() As String
    On Error Resume Next
'     CnnStrSQL = RetValue("HKLM\SOFTWARE\EAMS\CnnstrSQL")
     CnnStrSQL = "Provider=SQLOLEDB;Data Source=baxsgsinit02\baxsgsinit02;Initial Catalog=BAXPR_BAX_ASR;UID=asr.dbo;Pwd=baxasr;"
     
'     CnnStrSQL = "DSN=ASR;UID=asr.dbo;PWD=baxasr"
    'Provider=SQLOLEDB.1;Data Source=BAXSGSINBRIOR;Initial Catalog=DTMS;User ID=dtmsuser;Password=dtms;
End Property

Public Property Get MaxRecords() As Integer
    On Error Resume Next
    MaxRecords = RetValue("HKLM\SOFTWARE\EAMS\MaxRecordsNo")
    If MaxRecords = 0 Then MaxRecords = 300
End Property

Public Property Get SMTPHost() As String
    On Error Resume Next
    SMTPHost = RetValue("HKLM\SOFTWARE\EAMS\SMTPHost")
End Property

'Public Property Get HourRate() As Integer
'    On Error Resume Next
    'HourRate = RetValue("HKLM\SOFTWARE\EAMS\HourRate")
'    If HourRate = 0 Then HourRate = 40
'End Property

Public Function RetValue(strKey As String) As String
On Error Resume Next
    
'    "HKLM\SOFTWARE\EAMS\CnnstrSQL")
'
'    Dim obj As CRegObj
'
'    Set obj = New CRegObj
    
    Debug.Print GetSetting(appname:="SOFTWARE", section:="EAMS", Key:="CnnstrSQL", Default:="")
    RetValue = GetSetting(appname:="SOFTWARE", section:="EAMS", Key:="CnnstrSQL", Default:="")
'    RetValue = obj.Get(strKey)
'    Set obj = Nothing
End Function

Public Property Get ServerDate() As Date
On Error GoTo Err_Routine

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
          
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT GETDATE()"
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    
    Set rst = cmd.Execute
    ServerDate = rst.Fields(0)
    
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    
Exit_Routine:
    Screen.MousePointer = vbDefault
    Exit Property
    
Err_Routine:
    MsgBox Error$, vbInformation, "ServerDate"
    GoTo Exit_Routine
End Property

Public Function RecordExists(strSQL As String, ParamArray ipValues()) As Boolean

On Error GoTo Err_Routine
    
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim i As Integer
    
    RecordExists = False
    
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    
    For i = 0 To UBound(ipValues)
        cmd.Parameters(i).Value = ipValues(i)
    Next i
    
    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CnnStrSQL
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.Open cmd
    
    If rst.RecordCount > 0 Then RecordExists = True
    
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set rst = Nothing
    Set cmd = Nothing
    RecordExists = False
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function ReplaceNull(OrgValue, NewValue)
On Error GoTo Err_Routine
    
    If IsNull(OrgValue) Then
        ReplaceNull = NewValue
    Else
        ReplaceNull = OrgValue
    End If
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Function RetrieveList(strSQL As String, ParamArray ipValues()) As Recordset

On Error GoTo Err_Routine
    
    Dim cmd As ADODB.Command
    Dim i As Integer
      
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = CnnStrSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    
    For i = 0 To UBound(ipValues)
        cmd.Parameters(i).Value = ipValues(i)
    Next i
    
    Set RetrieveList = cmd.Execute
    
    Set cmd = Nothing
    
Exit_Routine:
    Exit Function
    
Err_Routine:
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine
    
End Function

Public Sub WriteLog(strMsg)
'Writing to a file
Dim iFileNum As Integer

    iFileNum = FreeFile
    Open App.Path & "\1msg.log" For Append As iFileNum
    Print #iFileNum, Now() & " " & strMsg
    Close iFileNum
End Sub

Public Function xSendMail(ByVal strTo As String, _
                         ByVal strCc As String, _
                         ByVal strSubject As String, _
                         ByVal strBody As String) As Long
                                                            

'SendMail = 1
'Exit Function
'send email
'Basic Error Handling
On Error GoTo EmailError

    Dim ns As Object
    Dim ndb As Object
    Dim ndoc As Object
    Dim strFrom As String
    Dim strToArr()      As String
    Dim strCCArr()      As String
    Dim lngStart        As Long
    Dim i               As Integer

strFrom = "ASR@baxglobal.com"
    
Set ns = CreateObject("Notes.Notessession")
Set ndb = ns.GETDATABASE("", "")
'set database to default
Call ndb.OPENMAIL

Set ndoc = ndb.CREATEDOCUMENT

lngStart = 1
i = 0
Do While InStr(lngStart, strTo, ",") > 0
    ReDim strToArr(i) As String
    lngStart = InStr(lngStart, strTo, ",") + 2
    i = i + 1
Loop
    
lngStart = 1
i = 0
Do While InStr(lngStart, strTo, ",") > 0
    strToArr(i) = Trim(Mid(strTo, lngStart, InStr(lngStart, strTo, ",") - lngStart))
    lngStart = InStr(lngStart, strTo, ",") + 2
    i = i + 1
Loop

lngStart = 1
i = 0
Do While InStr(lngStart, strCc, ",") > 0
    ReDim strCCArr(i) As String
    lngStart = InStr(lngStart, strCc, ",") + 2
    i = i + 1
Loop
    
lngStart = 1
i = 0
Do While InStr(lngStart, strCc, ",") > 0
    strCCArr(i) = Trim(Mid(strCc, lngStart, InStr(lngStart, strCc, ",") - lngStart))
    lngStart = InStr(lngStart, strCc, ",") + 2
    i = i + 1
Loop

Call ndoc.APPENDITEMVALUE("SendTo", strToArr)
Call ndoc.APPENDITEMVALUE("From", strFrom)
If strCc <> "" Then Call ndoc.APPENDITEMVALUE("CopyTo", strCCArr)
Call ndoc.APPENDITEMVALUE("Subject", strSubject)
Call ndoc.APPENDITEMVALUE("Body", strBody)

Call ndoc.Send(False)

xSendMail = 1

Set ns = Nothing
Set ndb = Nothing
Set ndoc = Nothing

EmailExit:
    Exit Function
  
EmailError:
    'Failed
    xSendMail = 0
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume EmailExit
End Function

'Public Function SendMail(ByVal strTo As String, _
'                         ByVal strCc As String, _
'                         ByVal strSubject As String, _
'                         ByVal strBody As String) As String
'On Error GoTo EmailError
'
'    Dim ObjMail As CDONTS.NewMail
'
'    Set ObjMail = New CDONTS.NewMail
'
'    ObjMail.From = "ASR@baxglobal.com"
'    ObjMail.To = strTo
'    If strCc <> "" Then ObjMail.Cc = strCc
'    ObjMail.Subject = strSubject
'    ObjMail.Body = strBody
'    ObjMail.Send
'    SendMail = "1"
'Set ObjMail = Nothing
'
'EmailExit:
'    Exit Function
'
'EmailError:
'    'Failed
'    SendMail = ""
'    Err.Raise Err.Number, Err.Source, Err.Description
'    Resume EmailExit
'
'End Function

'Using Persits ASP Mail Component
'Require installation of the Mail Component
Public Function SendMail(ByVal strTo As String, _
                         ByVal strCc As String, _
                         ByVal strSubject As String, _
                         ByVal strBody As String) As String
On Error GoTo EmailError
    Dim strHost As String
    Dim Mail As Object
    
    strHost = SMTPHost
    
    Set Mail = CreateObject("Persits.MailSender")
    
    ' enter valid SMTP host
    Mail.Host = strHost
    
    Mail.From = "ASR@schenker.com"
    Mail.FromName = "Schenker ASR"
    
    Mail.AddAddress strTo
    
    If strCc <> "" Then
        Mail.AddCC strCc
    End If
    ' message subject
    Mail.Subject = strSubject
    ' message body
    Mail.Body = strBody
    
    Mail.Send ' send message
    
    SendMail = "1"
    
    Set Mail = Nothing

EmailExit:
    Exit Function
  
EmailError:
    'Failed
    SendMail = ""
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume EmailExit

End Function

