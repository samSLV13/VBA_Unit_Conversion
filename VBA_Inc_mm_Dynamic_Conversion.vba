Sub Worksheet_Change(ByVal Target As Range)
    
    On Error GoTo ExitProcess 'If any error happens, ensures events are enabled again
    
    'Changed value debug print
    Debug.Print ("Target = " & Target)
    
    'Get Variable name from Target change
    Dim Variable_Name As String
    Variable_Name = Get_Last_Cell_Value(Target:=Target)
    Debug.Print ("Variable_Name = " & Variable_Name)

    If Trim(Variable_Name) = "" Then
        Debug.Print ("Early exit due to empty Variable_Name")
        Exit Sub
    End If

    Dim Current_Sheet_Name As String
    Dim Current_Sheet_Prefix  As String
    
    'Get the sheet name where changed was triggered
    Current_Sheet_Name = Get_Sheet_Name(ws:=Me)
    Debug.Print ("Current_Sheet_Name = " & Current_Sheet_Name)
    
    'Get the sheet prefix, will be used to group sheets with the same prefix
    Current_Sheet_Prefix = Get_Prefix_Sheet_Name(Sheet_Name:=Current_Sheet_Name)
    Debug.Print ("Current_Sheet_Prefix = " & Current_Sheet_Prefix)
    
    Debug.Print ("")
    
    Dim ws As Worksheet
    Dim Iterated_Sheet_Name As String
    Dim Iterated_Sheet_Prefix As String
    For Each ws In ThisWorkbook.Worksheets
        Iterated_Sheet_Name = Get_Sheet_Name(ws:=ws)
        Debug.Print ("Iterated_Sheet_Name = " & Iterated_Sheet_Name)
    
        Iterated_Sheet_Prefix = Get_Prefix_Sheet_Name(Sheet_Name:=Iterated_Sheet_Name)
        Debug.Print ("Iterated_Sheet_Prefix = " & Iterated_Sheet_Prefix)
        
        
        Debug.Print ("")
    Next ws
        
    
    'Application.EnableEvents = False 'Prevents code doesnt trigger itself when updating cells, prevents recursion
    '
    
    
    
    'Verify if value changed in C3 or E3 individually
    'If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
    '    If IsNumeric(Me.Range("C3").Value) Then
    '        Me.Range("E3").Value = Me.Range("C3").Value / 25.4 'Convert mm to in, then update E3 cell
    '    End If
    'ElseIf Not Intersect(Target, Me.Range("E3")) Is Nothing Then
    '    If IsNumeric(Me.Range("E3").Value) Then
    '        Me.Range("C3").Value = Me.Range("E3").Value * 25.4 'Convert in to mm, then update C3
    '    End If
    'End If
        
ExitProcess:
    Application.EnableEvents = True
        
End Sub

Function Get_Last_Cell_Value(Target As Range) As String
    Dim LastCell As Range
    Set LastCell = Target.End(xlToLeft)
    Get_Last_Cell_Value = LastCell.Value
End Function

Function Get_Sheet_Name(ws As Worksheet) As String
    Get_Sheet_Name = ws.Name
End Function

Function Get_Prefix_Sheet_Name(Sheet_Name As String) As String
    Dim Sheet_Name_Has_Underscore As Boolean
    Sheet_Name_Has_Underscore = Undescore_Exists_In_String(Sheet_Name:=Sheet_Name)
    If Sheet_Name_Has_Underscore = False Then
        Get_Prefix_Sheet_Name = "N/A"
    Else:
        Get_Prefix_Sheet_Name = Left(Sheet_Name, InStr(1, Sheet_Name, "_") - 1)
    End If
End Function

Function Undescore_Exists_In_String(Sheet_Name As String) As Boolean
    Undescore_Exists_In_String = True
    Dim UnderscorePosition As Integer
    UnderscorePosition = Get_Underscore_Position(Sheet_Name:=Sheet_Name)
    If UnderscorePosition = 0 Then Undescore_Exists_In_String = False
End Function

Function Get_Underscore_Position(Sheet_Name As String) As Integer
    Get_Underscore_Position = InStr(1, Sheet_Name, "_")
End Function
