Attribute VB_Name = "Module1"
Sub MacroMergeCells()
Attribute MacroMergeCells.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MacroMergeCells Macro
'

'
    'Range("A1:C1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub
Sub MacroUnmergeCells()
Attribute MacroUnmergeCells.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MacroUnmergeCells Macro
'

'
    'Range("A1:C1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
