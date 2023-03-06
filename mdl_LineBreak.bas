'Module: VBA Line Break


'#####
Public Sub Remove (strText as String) as String

    
    
    '##### Code #####
    If Len(strText) <> 0 Then
        If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then 
            strText = Left$(strText, Len(strText) - 2)
        End If
    End If
    Remove = strText
    
    
    
    '##### Optional #####
    'Remove = Application.WorksheetFunction.Clean(strText)
    
    
    
End Sub
