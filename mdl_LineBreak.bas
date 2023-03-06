'Module: VBA Line Break


'#####
Public Sub Remove (strText as String)

    If Len(strText) <> 0 Then
        If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then 
            strText = Left$(strText, Len(strText) - 2)
        End If
    End If

End Sub
