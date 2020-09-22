Attribute VB_Name = "modCpuName"
'modComputerName.bas

Option Explicit

Private Declare Function GetComputerName Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function ComputerName() As String
    Dim strBuffer As String
    Dim lngReturn As Long
    Dim strName As String
     
    strName = ""
    strBuffer = Space$(255)
    lngReturn = GetComputerName(strBuffer, 255)
      
    If lngReturn Then
        strName = Left$(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    End If
      
    ComputerName = strName
End Function

