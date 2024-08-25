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

    'Get Origin Value from Target change
    Dim Origin_Value As Double
    Origin_Value = Target
    Debug.Print ("Origin_Value = " & Origin_Value)

    'Get Conversion factor from Catalog Sheet
    Dim Conversion_Factor As Double
    Conversion_Factor = Get_Conversion_Factor(Origin_Unit:=Origin_Unit, Conversion_Unit:=Conversion_Unit)
    Debug.Print ("Conversion_Factor = " & Conversion_Factor)

    'Calculate Conversion Value
    Dim Conversion_Value As Double
    Conversion_Value = Origin_Value * Conversion_Factor
    Debug.Print ("Conversion_Value = " & Conversion_Value)

    'Update Conversion Value in Target change
    Target.Offset(0, 4).Value = Conversion_Value
    Debug.Print ("Conversion_Value updated in Target.Offset(0, 4) = " & Target.Offset(0, 4).Value)
    
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

Function Get_Conversion_Factor(Origin_Unit As String, Conversion_Unit As String) As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("UnitsCatalog")
    'Extract the catalog as an array
    Dim Catalog() As Variant
    Catalog = Extract_Array(Sheet_Name:="UnitsCatalog", Initial_Row:=1, Initial_Column:="A", Last_Column:="D", Number_Columns:=4)
    Debug.Print ("stop")
End Function

Function Extract_Array(Sheet_Name As String, Initial_Row As Long, Initial_Column As String, Last_Column As String, Number_Columns As Long) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(Sheet_Name)

    Dim Last_Cell As Long
    Last_Cell = ws.Cells(Initial_Row, Initial_Column).End(xlDown).Row

    Dim Aux_Array() As Variant
    ReDim Aux_Array(1 To Last_Cell, 1 To Number_Columns)
    Aux_Array = ws.Range(ws.Cells(Initial_Row, Initial_Column), ws.Cells(Last_Cell, Last_Column)).Value
    Extract_Array = Aux_Array
End Function
