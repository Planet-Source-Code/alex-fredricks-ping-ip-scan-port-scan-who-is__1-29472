Attribute VB_Name = "ModKeyCheck"
Option Explicit

Public Function KeyCheck(ByVal intKey As Integer, keyType As String) As Integer
 
    'global function to check keystrokes
    If keyType = "Alpha" Then
        
        Select Case intKey
            Case 65 To 90 'A to Z
                KeyCheck = intKey
            Case 97 To 122 'a to z
                KeyCheck = intKey
            Case vbKey0 To vbKey9
                KeyCheck = intKey
            Case vbKeyBack
                KeyCheck = intKey
            Case vbKeySpace
                KeyCheck = intKey
            Case Asc(".")
                KeyCheck = intKey
            Case Else
                KeyCheck = 0
        End Select
    
    ElseIf keyType = "Num" Then
        
        Select Case intKey
            Case Asc("0") To Asc("9")
                KeyCheck = intKey
            Case vbKeyBack
                KeyCheck = intKey
            Case Asc(".")
                KeyCheck = 0
            Case Else
                KeyCheck = 0
        End Select
    End If

End Function





