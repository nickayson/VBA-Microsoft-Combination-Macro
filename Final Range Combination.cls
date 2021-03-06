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
    'Check if a Destination sheet exists with function call, creates a new one if false
    If SheetExists("Results") = False Then
        Sheets.Add
        ActiveSheet.Name = "Results"
    End If
    
    'Variable declarations ranges and format rule
    Dim rng As Range
    Dim xCell As Range
        
    Dim DefaultRange As Range
    Dim FormatRuleInput As String
        
    Dim pasteSheet As Worksheet
    Set pasteSheet = Worksheets("Results")

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
    
    'Replace all empty cells with a space just for formatting purposes
    Application.ScreenUpdating = False
    For Each xCell In rng
        If Trim(xCell.Value) = "" Then
            xCell.Value = " "
        End If
    Next
    Application.ScreenUpdating = True
    
    'Copy range of cells selected and print directly underneath existing data in pastesheet
    Application.ScreenUpdating = False
    rng.Copy
    pasteSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
        
End Sub

'Function to check if new destination file exists

Function SheetExists(shName As String) As Boolean
    With ActiveWorkbook
        On Error Resume Next
        SheetExists = (.Sheets(shName).Name = shName)
        On Error GoTo 0
    End With
    
End Function

