Attribute VB_Name = "modFocus"
Option Explicit

Public Sub Focus(varX As Variant)
'selects entire txtbox
    With varX
        If .Text <> "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

