'Module: VBA Line Break


'#####
Public Sub RemoveFromString (strText as String) as String
    
    '##### Code #####
    If Len(strText) <> 0 Then
        If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then 
            strText = Left$(strText, Len(strText) - 2)
        End If
    End If
    RemoveFromString = strText
    
    
    '##### Optional #####
    'RemoveFromString = Application.WorksheetFunction.Clean(strText)
    

'#####
End Sub
