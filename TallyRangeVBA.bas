Attribute VB_Name = "Module1"
'''''' Run TallyRange()
Function TallyRange(arg)

    arg.Parent.Evaluate ("0+UniqueStrings(" & arg.Address(False, False) & ",""" & Application.Caller.Worksheet.Name & """," & Application.Caller.Row & "," & Application.Caller.Column & ")")
                '0+ is placed above to workaround the "evaluating twice" bug in VBA
                
    TallyRange = "Range(" & arg.Address(False, False) & ") tallied"  'or whatever return value is useful...

End Function


Sub UniqueStrings(args As Range, sheetName As String, origin_r As Integer, origin_c As Integer)
Attribute UniqueStrings.VB_ProcData.VB_Invoke_Func = " \n14"
     
    'Columns for output
    Dim name_c As Integer
    name_c = origin_c

    Dim cnt_c As Integer
    cnt_c = origin_c + 1

    'Bottom row of output list
    Dim originListEnd_r As Integer
    originListEnd_r = origin_r + 1
    
    'Loop over all cells in range
    For Each cell In args
        Dim temp_r As Integer
        temp_r = findValueRow(sheetName, origin_r, originListEnd_r, name_c, cell.Value)
        If temp_r = -1 Then 'If string was not found in args
            Worksheets(sheetName).Cells(originListEnd_r, name_c).Value = cell.Value
            Worksheets(sheetName).Cells(originListEnd_r, cnt_c).Value = 1
            originListEnd_r = originListEnd_r + 1
        Else 'String was found
            Worksheets(sheetName).Cells(temp_r, cnt_c).Value = Worksheets(sheetName).Cells(temp_r, cnt_c).Value + 1
            
        End If

    Next cell

    'Dashes to denote ending data
    Worksheets(sheetName).Cells(originListEnd_r, name_c).Value = "-"
    Worksheets(sheetName).Cells(originListEnd_r, cnt_c).Value = "-"
    
End Sub

Function findValueRow(sheetName As String, start_r As Integer, end_r As Integer, col_c As Integer, val As String) As Integer
    Dim r As Integer
    r = start_r
    
    Do While r <> end_r
        If Worksheets(sheetName).Cells(r, col_c).Value = val Then
            findValueRow = r
            Exit Function
        End If
        r = r + 1
    Loop

    findValueRow = -1 'Not Found value

End Function

