Attribute VB_Name = "modUserName"
'modUserName.bas

Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function UserName() As String
    Dim lngReturn    As Long
    Dim strUserName  As String
    Dim strBuffer    As String
        
    strUserName = ""
    strBuffer = Space$(255)
    lngReturn = GetUserName(strBuffer, 255)
        
    If lngReturn Then
    strUserName = Left$(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    End If
        
    UserName = strUserName
End Function


