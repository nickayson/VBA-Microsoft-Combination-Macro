VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub CopyFromMultiSheets()
Attribute CopyFromMultiSheets.VB_ProcData.VB_Invoke_Func = "k\n14"

    'Check if a Destination sheet exists with function call, creates a new one if false
    If SheetExists("Overall Summary") = False Then
        Sheets.Add
        ActiveSheet.name = "Overall Summary"
    End If
    
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim FirstSh As Worksheet
    Dim LastRow As Long
    Dim CopyRng As Range

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        
        'Destination sheet as results
        Set DestSh = Sheets("Overall Summary")
        'FirstSheet is detected at first sheet in the workbook
        Set FirstSh = Sheets(1)
        
        'loop through all worksheets in workbook
        For Each sh In ActiveWorkbook.Worksheets
            'Case for each possible outcome
            Select Case sh.name
            
            'Skip Destination Sheet if it is first sheet in workbook
            Case DestSh.name
            If DestSh.name = FirstSh.name Then
                Set FirstSh = Sheets(2)
            End If
            
            'Will only run 1 time to copy the first header of the first sheet
            Case FirstSh.name
            'Copy from worksheet until a cell in the column A is empty
            'It will then paste to Results sheet right after existing data
                LastRow = DestSh.Cells(Rows.Count, 1).End(3).Row + 1
                sh.Range("A5", sh.Range("N10").End(xlDown)).Copy
                Sheets("Overall Summary").Cells(LastRow, "A").PasteSpecial xlAll
                
            'If its not any of the cases listed above it will go through every sheet
            Case Else
            'Copy from worksheet from A6 to end of column A (No copying of headers)
            'Paste to Results sheet right after existing data
                If Not sh.name Like "s*" Then
                    LastRow = DestSh.Cells(Rows.Count, 1).End(3).Row + 1
                    sh.Range("A6", sh.Range("N10").End(xlDown)).Copy
                    Sheets("Overall Summary").Cells(LastRow, "A").PasteSpecial xlAll
                End If
                
            End Select
            
        'goes to next sheet
        Next sh
        
        'delete first row for offset formatting
        Rows(1).EntireRow.Delete

        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

'Function to check if new destination file exists
Function SheetExists(shName As String) As Boolean

    With ActiveWorkbook
        On Error Resume Next
        SheetExists = (.Sheets(shName).name = shName)
        On Error GoTo 0
    End With

End Function

Private Sub Workbook_Open()

End Sub
