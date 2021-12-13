Attribute VB_Name = "modTackleCombinations"
Sub CommandButton5_Click()

Dim rowCount As Integer: rowCount = 1

Do Until Range("Tackles_Selections_1").Cells(rowCount, 1) = ""

Select Case Range("Tackles_Selection_Count").Cells(rowCount, 1)

    Case "6"
        Range("Tackles_Selection_Names").Cells(rowCount, 1) = Range("Tackles_Selections_1").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_2").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_3").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_4").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_5").Cells(rowCount, 1) & " and " & Range("Tackles_Selections_6").Cells(rowCount, 1) & " to make " & Range("Tackles_Combinations").Cells(rowCount, 1) & " tackles between them"
    Case "5"
        Range("Tackles_Selection_Names").Cells(rowCount, 1) = Range("Tackles_Selections_1").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_2").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_3").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_4").Cells(rowCount, 1) & " and " & Range("Tackles_Selections_5").Cells(rowCount, 1) & " to make " & Range("Tackles_Combinations").Cells(rowCount, 1) & " tackles between them"
    Case "4"
        Range("Tackles_Selection_Names").Cells(rowCount, 1) = Range("Tackles_Selections_1").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_2").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_3").Cells(rowCount, 1) & " and " & Range("Tackles_Selections_4").Cells(rowCount, 1) & " to make " & Range("Tackles_Combinations").Cells(rowCount, 1) & " tackles between them"
    Case "3"
        Range("Tackles_Selection_Names").Cells(rowCount, 1) = Range("Tackles_Selections_1").Cells(rowCount, 1) & ", " & Range("Tackles_Selections_2").Cells(rowCount, 1) & " and " & Range("Tackles_Selections_3").Cells(rowCount, 1) & " to make " & Range("Tackles_Combinations").Cells(rowCount, 1) & " tackles between them"
    Case "2"
        Range("Tackles_Selection_Names").Cells(rowCount, 1) = Range("Tackles_Selections_1").Cells(rowCount, 1) & " and " & Range("Tackles_Selections_2").Cells(rowCount, 1) & " to make " & Range("Tackles_Combinations").Cells(rowCount, 1) & " tackles between them"
    Case Else
        Range("Tackles_Selection_Names").Cells(rowCount, 1) = ""
    
    End Select
    
50  rowCount = rowCount + 1
    
    Loop
    
    Call modTackleCombinations.ProtectTacklesSelections
    
End Sub

Public Sub ProtectTacklesSelections()

Dim selectionRow As Integer: selectionRow = 1

ThisWorkbook.Worksheets("Tackles Selections").Unprotect
Application.ScreenUpdating = False

Do Until selectionRow = 50

    If Range("Tackles_Selection_Names").Cells(selectionRow, 1) <> "" Then
    
        Range("Tackles_Selections_1").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Selections_2").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Selections_3").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Selections_4").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Selections_5").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Selections_6").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Combinations").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_True_Prices").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Offer_Prices").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Tackles_Selection_Names").Cells(selectionRow, 1).Interior.ColorIndex = 22

        Range("Tackles_Selections_1").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Selections_2").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Selections_3").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Selections_4").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Selections_5").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Selections_6").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Combinations").Cells(selectionRow, 1).Locked = True
        Range("Tackles_True_Prices").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Offer_Prices").Cells(selectionRow, 1).Locked = True
        Range("Tackles_Selection_Names").Cells(selectionRow, 1).Locked = True
        
    Else
       
    End If
    
    selectionRow = selectionRow + 1
    
Loop

Application.ScreenUpdating = True
ThisWorkbook.Worksheets("Shots Selections").Protect

End Sub

Public Sub tackleErrors()

Dim rowCount As Integer: rowCount = 1
Dim totalTackles As String
Dim selectionName As String

Do Until Range("Tackles_Combinations").Cells(rowCount, 1) = ""

    totalTackles = Range("Tackles_Combinations").Cells(rowCount, 1)
    selectionName = Range("Tackles_Selection_Names").Cells(rowCount, 1)
    
    If StringMatch(selectionName, totalTackles) = False Then
        MsgBox "Error with selection " & rowCount & " of tackle combinations!"
        End
    Else
      rowCount = rowCount + 1
    End If
    
    Loop

End Sub
