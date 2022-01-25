VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CombineWorksheets()

    'Ask user for name of sheet to copy from
    Dim inputsheet As Variant
    inputsheet = InputBox("Give me the name of the sheet you want to copy from:")
    
    'Ask user for name of sheet to send data to
    'Dim destsheet As Variant
    'destsheet = InputBox("Give me the name of the sheet you want to paste too:")
    'On Error GoTo 0
    
    'Check if a Destination sheet exists with function call, creates a new one if false
    If SheetExists("Results") = False Then
        Sheets.Add
        ActiveSheet.Name = "Results"
    End If
    
    'Chooses which worksheet you are working with
    With ThisWorkbook.Worksheets(inputsheet)

        'Variable declarations ranges and format rule
        Dim rng As Range
        Dim re As Range
        'Set rng = .Range("O66")

        Dim DefaultRange As Range
        Dim FormatRuleInput As String

        'Determine a default range based on user's Selection
        If TypeName(Selection) = "Range" Then
            Set DefaultRange = Selection
        Else
            Set DefaultRange = ActiveCell
        End If

        'Get A Cell Address From The User to Get Number Format From
        On Error Resume Next
        Set rng = Application.InputBox( _
            Title:="Find range of cells", _
            Prompt:="Select a cell range to copy from and paste to" & vbCrLf & "Use the format *Cellone:CellTwo*, Ex: A1:A2", _
            Default:=DefaultRange.Address, _
            Type:=8)
        On Error GoTo 0

        'Test to ensure User Did not cancel
        If rng Is Nothing Then Exit Sub
        
        Dim pasteSheet As Worksheet
        Set pasteSheet = Worksheets("Results")

    
        Application.ScreenUpdating = False
        rng.Copy
        pasteSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues, SkipBlanks:=False
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
        
  
    End With
End Sub

Sub CommandButton1_Click()
  Application.ScreenUpdating = False
  Dim copySheet As Worksheet
  Dim pasteSheet As Worksheet
  Dim rw As Integer

  Set copySheet = Worksheets("FPI_GPR-1114")
  Set pasteSheet = Worksheets("Results")
  
  copySheet.Range("A7:O65").Copy
  pasteSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
  Application.CutCopyMode = False
  Application.ScreenUpdating = True
  
End Sub
Sub Range_End_Method()
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 15).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
            "Last Column: " & lCol
  
End Sub
Sub GoToLastRowofRange()
   Dim rw As Integer
   Range("B1").Select
'get the last row in the current region
   rw = Range("B1").End(xlDown).Row
'show how many rows are used
   MsgBox "The last row used in this range is " & rw
End Sub

'Function to check if new destination file exists

Function SheetExists(shName As String) As Boolean
    With ActiveWorkbook
        On Error Resume Next
        SheetExists = (.Sheets(shName).Name = shName)
        On Error GoTo 0
    End With
    
End Function

Private Sub Workbook_Open()

End Sub

