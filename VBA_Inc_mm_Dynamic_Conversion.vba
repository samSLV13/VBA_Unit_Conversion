Sub Worksheet_Change(ByVal Target As Range)
    
    'If any error happens, ensures events are enabled again
    On Error GoTo ExitProcess
    
    'Changed value debug print
    Debug.Print ("Target = " & Target)
    
    'Get Variable name from Target change
    Dim Variable_Name As String
    Variable_Name = Get_Last_Cell_Value(Target:=Target)
    Debug.Print ("Variable_Name = " & Variable_Name)

    'Exit early if the row of the changed value doesnt have a string to assign as variable_name
    If Trim(Variable_Name) = "" Then Debug.Print ("Early exit due to empty Variable_Name"): Exit Sub

    'Get Origin Unit from Target change
    Dim Origin_Unit As String
    Origin_Unit = Target.Offset(0, 1).Value
    Debug.Print ("Origin_Unit = " & Origin_Unit)

    'Get Conversion Unit from Target change
    Dim Conversion_Unit As String
    Conversion_Unit = Target.Offset(0, 3).Value
    Debug.Print ("Conversion_Unit = " & Conversion_Unit)

    'Get the sheet name where changed was triggered
    Dim Current_Sheet_Name As String
    Current_Sheet_Name = Get_Sheet_Name(ws:=Me)
    Debug.Print ("Current_Sheet_Name = " & Current_Sheet_Name)
    
    'Get the sheet prefix, will be used to group sheets with the same prefix
    Dim Current_Sheet_Prefix  As String
    Current_Sheet_Prefix = Get_Prefix_Sheet_Name(Sheet_Name:=Current_Sheet_Name)
    Debug.Print ("Current_Sheet_Prefix = " & Current_Sheet_Prefix)
    
    Debug.Print ("")
    
    'Process value in each sheet from Workbook
    Debug.Print ("Start iterating sheets...")
    Debug.Print ("")
    Dim ws As Worksheet
    Dim Iterated_Sheet_Name As String, Iterated_Sheet_Prefix As String
    For Each ws In ThisWorkbook.Worksheets
        Iterated_Sheet_Name = Get_Sheet_Name(ws:=ws)
        Debug.Print ("Iterated_Sheet_Name = " & Iterated_Sheet_Name)
    
        Iterated_Sheet_Prefix = Get_Prefix_Sheet_Name(Sheet_Name:=Iterated_Sheet_Name)
        Debug.Print ("Iterated_Sheet_Prefix = " & Iterated_Sheet_Prefix)
        
        If Iterated_Sheet_Prefix <> Current_Sheet_Prefix Then GoTo Next_Iteration

        Dim Found_Cell As Range
        Set Found_Cell = ws.Range("A1:Z100").Find(What:=Variable_Name, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If Found_Cell Is Nothing Then GoTo Next_Iteration
        Debug.Print ("Found_Cell = " & Found_Cell.Address)

        Dim Found_Cell_Value As String
        Dim Found_Cell_Column As Integer
        Found_Cell_Column = Found_Cell.Column
        Found_Cell_Value = ws.Cells(Found_Cell.Row, Found_Cell_Column + 1).Value
        Debug.Print ("Found_Cell_Value = " & Found_Cell_Value)

Next_Iteration:

        Debug.Print ("")
    Next ws
        
    
    '
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

