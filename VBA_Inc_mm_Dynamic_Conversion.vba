Sub Worksheet_Change(ByVal Target As Range)
    
    'If Target value is a string, exit early
    If Not IsNumeric(Target.Value) Then
        Debug.Print ("Early exit, Value is not a number")
        Exit Sub
    End If
    
    'If any error happens, ensures events are enabled again
    Application.EnableEvents = False
    On Error GoTo ExitProcess
    
    'Changed value debug print
    Debug.Print ("Target = " & Target)
    
    'Get Sheet Data
    Dim Current_Sheet_Name As String, Current_Sheet_Prefix  As String
    Call Get_Sheet_Data(Target:=Target, Current_Sheet_Name:=Current_Sheet_Name, Current_Sheet_Prefix:=Current_Sheet_Prefix)
    
    'Get Initial Values
    Dim Variable_Name As String, Origin_Value As Double
    Call Get_Initial_Values(Target:=Target, Variable_Name:=Variable_Name, Origin_Value:=Origin_Value)

    'Exit early if the row of the changed value doesnt have a string to assign as variable_name
    If Trim(Variable_Name) = "" Then
        Debug.Print ("Early exit due to empty Variable_Name")
        Application.EnableEvents = True
        Exit Sub
    End If

    'Get Conversion Values
    Dim Origin_Unit As String, Conversion_Unit As String, Conversion_Factor As Double, Conversion_Operation As String
    Call Get_Conversion_Values(Target:=Target, Origin_Unit:=Origin_Unit, Conversion_Unit:=Conversion_Unit, Conversion_Factor:=Conversion_Factor, Conversion_Operation:=Conversion_Operation)

    'Get Converted Value
    Dim Converted_Value As Double
    Call Get_Converted_Value(Origin_Value:=Origin_Value, Conversion_Factor:=Conversion_Factor, Conversion_Operation:=Conversion_Operation, Converted_Value:=Converted_Value)

    'Update Conversion Value in Target change
    Call Update_Converted_Value(Target:=Target, Converted_Value:=Converted_Value, Variable_Name:=Variable_Name)
    
    'Process value in each sheet from Workbook
    Debug.Print ("")
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
        Set Found_Cell = ws.Range("A1:Z100").Find(What:=Variable_Name, LookIn:=xlValues, LookAt:=xlWhole)
        If Found_Cell Is Nothing Then GoTo Next_Iteration
        Debug.Print ("Found_Cell = " & Found_Cell.Address)

        'Update Origin Value, 1 columna after Found Cell
        Found_Cell.Offset(0, 1).Value = Origin_Value

        'Update Converted Value, 3 columnas after Found Cell
        Found_Cell.Offset(0, 3).Value = Converted_Value

Next_Iteration:

        Debug.Print ("")
    Next ws

ExitProcess:
    Application.EnableEvents = True
        
End Sub

Function Undescore_Exists_In_String(Sheet_Name As String) As Boolean
    Undescore_Exists_In_String = True
    Dim UnderscorePosition As Integer
    UnderscorePosition = Get_Underscore_Position(Sheet_Name:=Sheet_Name)
    If UnderscorePosition = 0 Then Undescore_Exists_In_String = False
End Function

Function Get_Underscore_Position(Sheet_Name As String) As Integer
    Get_Underscore_Position = InStr(1, Sheet_Name, "_")
End Function

Function Get_Last_Cell_Value(Target As Range) As String
    Dim LastCell As Range
    Set LastCell = Target.End(xlToLeft)
    Get_Last_Cell_Value = LastCell.Value
End Function

Sub Get_Sheet_Data(Target As Range, ByRef Current_Sheet_Name As String, ByRef Current_Sheet_Prefix As String)
    Debug.Print ("")
    Debug.Print ("Obtaining Sheet Data...")
    Current_Sheet_Name = Get_Sheet_Name(ws:=Me)
    Current_Sheet_Prefix = Get_Prefix_Sheet_Name(Sheet_Name:=Current_Sheet_Name)
    Debug.Print ("Current_Sheet_Name = " & Current_Sheet_Name)
    Debug.Print ("Current_Sheet_Prefix = " & Current_Sheet_Prefix)
End Sub

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

Sub Get_Initial_Values(Target As Range, ByRef Variable_Name As String, ByRef Origin_Value As Double)
    Debug.Print ("")
    Debug.Print ("Obtaining Initial Values...")
    Variable_Name = Get_Last_Cell_Value(Target:=Target)
    Origin_Value = Target
    Debug.Print ("Variable_Name = " & Variable_Name)
    Debug.Print ("Origin_Value = " & Origin_Value)
End Sub

Sub Get_Conversion_Values(Target As Range, ByRef Origin_Unit As String, ByRef Conversion_Unit As String, ByRef Conversion_Factor As Double, ByRef Conversion_Operation As String)
    Debug.Print ("")
    Debug.Print ("Obtaining Conversion Values...")
    Origin_Unit = Target.Offset(0, 1).Value
    Conversion_Unit = Target.Offset(0, 3).Value
    Dim Conversion_Config As Variant
    Conversion_Config = Extract_Array(Sheet_Name:="UnitsCatalog", Initial_Row:=1, Initial_Column:="A", Last_Column:="D", Number_Columns:=4)
    Conversion_Factor = Get_Conversion_Factor(Origin_Unit:=Origin_Unit, Conversion_Unit:=Conversion_Unit, Conversion_Config:=Conversion_Config)
    Conversion_Operation = Get_Conversion_Operation(Origin_Unit:=Origin_Unit, Conversion_Unit:=Conversion_Unit, Conversion_Config:=Conversion_Config)
    Debug.Print ("Origin_Unit = " & Origin_Unit)
    Debug.Print ("Conversion_Unit = " & Conversion_Unit)
    Debug.Print ("Conversion_Factor = " & Conversion_Factor)
    Debug.Print ("Conversion_Operation = " & Conversion_Operation)
End Sub

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

Function Get_Conversion_Factor(Origin_Unit As String, Conversion_Unit As String, Conversion_Config As Variant) As Double
    Dim i As Long
    For i = LBound(Conversion_Config) To UBound(Conversion_Config)
        If Conversion_Config(i, 1) = Origin_Unit And Conversion_Config(i, 2) = Conversion_Unit Then Get_Conversion_Factor = Conversion_Config(i, 4): Exit Function
    Next i
End Function

Function Get_Conversion_Operation(Origin_Unit As String, Conversion_Unit As String, Conversion_Config As Variant) As String
    Dim i As Long
    For i = LBound(Conversion_Config) To UBound(Conversion_Config)
        If Conversion_Config(i, 1) = Origin_Unit And Conversion_Config(i, 2) = Conversion_Unit Then Get_Conversion_Operation = Conversion_Config(i, 3): Exit Function
    Next i
End Function

Sub Get_Converted_Value(Origin_Value As Double, Conversion_Factor As Double, Conversion_Operation As String, ByRef Converted_Value As Double)
    Debug.Print ("")
    Debug.Print ("Obtaining Converted Value...")
    If Conversion_Operation = "Multiply" Then
        Converted_Value = Origin_Value * Conversion_Factor
    ElseIf Conversion_Operation = "Divide" Then
        Converted_Value = Origin_Value / Conversion_Factor
    End If
    Debug.Print ("Converted_Value = " & Converted_Value)
End Sub

Sub Update_Converted_Value(Target As Range, Converted_Value As Double, Variable_Name As String)
    Debug.Print ("")
    Debug.Print ("Updating Converted Value...")
    Dim Converted_Cell As Range
    Set Converted_Cell = Target.Offset(0, 2)
    Converted_Cell.Value = Converted_Value
    Debug.Print ("Converted_Value updated in Converted_Cell = " & Converted_Cell.Value)
End Sub


