VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'USER_INFO_3 structure is defined in Win32 SDK, below is the VB
      'declare.
      'LPWSTR is a pointer to a Unicode string.
      'The storage of the USER_INFO_3  structure, including the string
      'buffer referred by its LPWSTR members, are all allocated by Windows
      'NT so the function caller doesn't have to allocate memory.


Private Type USER_INFO_3
    usri3_name As Long           'LPWSTR in SDK
    usri3_password As Long       'LPWSTR in SDK
    usri3_password_age As Long      'DWORD in SDK
    usri3_priv As Long           'DWORD in SDK
    usri3_home_dir As Long       'LPWSTR in SDK
    usri3_comment As Long        'LPWSTR in SDK
    usri3_flags As Long          'DWORD in SDK
    usri3_script_path As Long    'LPWSTR in SDK
    usri3_auth_flags As Long        'DWORD in SDK
    usri3_full_name As Long         'LPWSTR in SDK
    usri3_usr_comment As Long    'LPWSTR in SDK
    usri3_parms As Long          'LPWSTR in SDK
    usri3_workstations As Long      'LPWSTR in SDK
    usri3_last_logon As Long        'DWORD in SDK
    usri3_last_logoff As Long    'DWORD in SDK
    usri3_acct_expires As Long      'DWORD in SDK
    usri3_max_storage As Long    'DWORD in SDK
    usri3_units_per_week As Long    'DWORD in SDK
    usri3_logon_hours As Long    'PBYTE in SDK
    usri3_bad_pw_count As Long      'DWORD in SDK
    usri3_num_logons As Long        'DWORD in SDK
    usri3_logon_server As Long      'LPWSTR in SDK
    usri3_country_code As Long      'DWORD in SDK
    usri3_code_page As Long         'DWORD in SDK
    usri3_user_id As Long        'DWORD in SDK
    usri3_primary_group_id As Long  'DWORD in SDK
    usri3_profile As Long        'LPWSTR in SDK
    usri3_home_dir_drive As Long    'LPWSTR in SDK
    usri3_password_expired As Long  'DWORD in SDK
End Type

Private Declare Function NetUserGetInfo Lib "netapi32.dll" ( _
   strServerName As Any, strUserName As Any, ByVal dwLevel As Long, _
   pBuffer As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Const NERR_SUCCESS = 0

Public Function RetrieveUser(ByVal strId As String) As String
     
On Error GoTo Err_Routine

    Dim pServer() As Byte, pUser() As Byte
    Dim dwLevel As Long
    Dim tmpBuffer As USER_INFO_3
    Dim ptmpBuffer As Long
    Dim sUser As String
    Dim lngRet As Long
    
    pUser = strId & vbNullChar
    pServer = "" & vbNullChar
              
    dwLevel = 3
    
    lngRet = NetUserGetInfo(pServer(0), pUser(0), dwLevel, ptmpBuffer)
    If (lngRet = NERR_SUCCESS) Then
        CopyMemory tmpBuffer, ptmpBuffer, LenB(tmpBuffer)
        
        Dim sByte() As Byte
        ReDim sByte(255)
        CopyMemory sByte(0), tmpBuffer.usri3_full_name, 256
        sUser = sByte
    Else
        sUser = strId
    End If
    'sUser = sUser & vbNullChar
    RetrieveUser = Trim(sUser)

Exit_Routine:
    Exit Function
    
Err_Routine:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Exit_Routine

End Function


