Attribute VB_Name = "saveCSV"
Sub saveCSV()
Dim SavingPath As String
Dim UniqueNameBit As String

SavingPath = "X:\Bet Tribe\Trading\Risk Uploads\Football\OPTA Combined" 'this file path needs to be changed if you want to save elsewhere

Application.ScreenUpdating = False
Application.DisplayAlerts = False

ActiveWorkbook.Sheets("2").Visible = True
ActiveWorkbook.Sheets("3").Visible = True
    
    
    Sheets("2").Activate
    Sheets("2").Cells.Select
    Selection.AutoFilter Field:=38, Criteria1:="Export" 'export criteria can be changed in the '2' sheet
       Range("A1").Select
    Sheets("2").Range(Selection, Selection.End(xlDown)).Select
    Sheets("2").Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ChDir "C:\"
    ActiveWorkbook.SaveAs Filename:="C:\CSV\CombinedMarkets.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    
    'change the 'CombinedMarkets' naming here if you want to rename how the CSV saves

    Sheets("3").Activate
    Sheets("3").Cells.Select
    Selection.AutoFilter Field:=29, Criteria1:="Export" 'export criteria can be changed in the '3' sheet
       Range("A1").Select
    Sheets("3").Range(Selection, Selection.End(xlDown)).Select
    Sheets("3").Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Columns("AC:AC").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Application.CutCopyMode = False
    ChDir "C:\"
    ActiveWorkbook.SaveAs Filename:="C:\CSV\CombinedSelections.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    
     'change the 'CombinedSelections' naming here if you want to rename how the CSV saves
    
    
    Sheets("3").Activate
    With Sheets("3")
    .AutoFilterMode = False
    End With
    
    Range("A1").Select

ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1

Application.ScreenUpdating = True
Application.DisplayAlerts = True

ActiveWorkbook.Sheets("2").Visible = False
ActiveWorkbook.Sheets("3").Visible = False

MsgBox ("CSVs have been saved")

Sheets("Match Setup").Activate

Dim saveString As String
        saveString = "X:\Bet Tribe\Trading\Risk Uploads\Football\OPTA Combined\" & Sheets("Match Setup").Range("V3") & ".xlsm" 'reference to cell V3 is the event name - make sure this stays in that cell or the file will save blank
        ActiveWorkbook.SaveAs Filename:=saveString

MsgBox ("Saved To Risk Uploads")

End Sub

Sub SaveResults()

Sheets("Results CSV").Visible = True
Sheets("Results CSV").Activate
Sheets("Results CSV").Cells.Select
Selection.AutoFilter Field:=13, Criteria1:="Export"

Sheets("Results CSV").Range("A1").Select
Sheets("Results CSV").Range(Selection, Selection.End(xlDown)).Select
Sheets("Results CSV").Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Workbooks.Add
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False

Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    
Range("A1").Select
Application.CutCopyMode = False
ChDir "C:\"
ActiveWorkbook.SaveAs Filename:="C:\CSV\OptaCombinedResults.csv", FileFormat:=xlCSV, _
    CreateBackup:=False
ActiveWindow.Close
Sheets("Results CSV").Visible = False
Sheets("Resulting Inputs").Visible = False
Sheets("Match Setup").Activate

Application.ScreenUpdating = True

MsgBox ("ResultsCSV Saved" & vbNewLine & "Upload 'OptaCombinedResults' to Admin")

Dim saveString As String
    saveString = "X:\Bet Tribe\Trading\Risk Uploads\Football\OPTA Combined\Resulting\" & Sheets("Match Setup").Range("V3") & " - Results" & ".xlsm"
    ActiveWorkbook.SaveAs Filename:=saveString
    
End Sub

