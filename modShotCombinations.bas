Attribute VB_Name = "modShotCombinations"
Public Sub CommandButton3_Click()

Dim rowCount As Integer: rowCount = 1

Do Until Range("Shots_Selections_1").Cells(rowCount, 1) = ""

Select Case Range("Shots_Selection_Count").Cells(rowCount, 1)

    Case "6"
        Range("Shots_Selection_Names").Cells(rowCount, 1) = Range("Shots_Selections_1").Cells(rowCount, 1) & ", " & Range("Shots_Selections_2").Cells(rowCount, 1) & ", " & Range("Shots_Selections_3").Cells(rowCount, 1) & ", " & Range("Shots_Selections_4").Cells(rowCount, 1) & ", " & Range("Shots_Selections_5").Cells(rowCount, 1) & " and " & Range("Shots_Selections_6").Cells(rowCount, 1) & " to have " & Range("Shots_Combinations").Cells(rowCount, 1) & " shots on target between them"
    Case "5"
        Range("Shots_Selection_Names").Cells(rowCount, 1) = Range("Shots_Selections_1").Cells(rowCount, 1) & ", " & Range("Shots_Selections_2").Cells(rowCount, 1) & ", " & Range("Shots_Selections_3").Cells(rowCount, 1) & ", " & Range("Shots_Selections_4").Cells(rowCount, 1) & " and " & Range("Shots_Selections_5").Cells(rowCount, 1) & " to have " & Range("Shots_Combinations").Cells(rowCount, 1) & " shots on target between them"
    Case "4"
        Range("Shots_Selection_Names").Cells(rowCount, 1) = Range("Shots_Selections_1").Cells(rowCount, 1) & ", " & Range("Shots_Selections_2").Cells(rowCount, 1) & ", " & Range("Shots_Selections_3").Cells(rowCount, 1) & " and " & Range("Shots_Selections_4").Cells(rowCount, 1) & " to have " & Range("Shots_Combinations").Cells(rowCount, 1) & " shots on target between them"
    Case "3"
        Range("Shots_Selection_Names").Cells(rowCount, 1) = Range("Shots_Selections_1").Cells(rowCount, 1) & ", " & Range("Shots_Selections_2").Cells(rowCount, 1) & " and " & Range("Shots_Selections_3").Cells(rowCount, 1) & " to have " & Range("Shots_Combinations").Cells(rowCount, 1) & " shots on target between them"
    Case "2"
        Range("Shots_Selection_Names").Cells(rowCount, 1) = Range("Shots_Selections_1").Cells(rowCount, 1) & " and " & Range("Shots_Selections_2").Cells(rowCount, 1) & " to have " & Range("Shots_Combinations").Cells(rowCount, 1) & " shots on target between them"
    Case Else
        Range("Shots_Selection_Names").Cells(rowCount, 1) = ""
    
    End Select
    
50  rowCount = rowCount + 1
    
    Loop
    
    Call modShotCombinations.ProtectShotsSelections
    
End Sub

Public Sub ProtectShotsSelections()

Dim selectionRow As Integer: selectionRow = 1

ThisWorkbook.Worksheets("Shots Selections").Unprotect
Application.ScreenUpdating = False

Do Until selectionRow = 50

    If Range("Shots_Selection_Names").Cells(selectionRow, 1) <> "" Then
    
        Range("Shots_Selections_1").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Selections_2").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Selections_3").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Selections_4").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Selections_5").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Selections_6").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Combinations").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_True_Prices").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Offer_Prices").Cells(selectionRow, 1).Interior.ColorIndex = 22
        Range("Shots_Selection_Names").Cells(selectionRow, 1).Interior.ColorIndex = 22

        Range("Shots_Selections_1").Cells(selectionRow, 1).Locked = True
        Range("Shots_Selections_2").Cells(selectionRow, 1).Locked = True
        Range("Shots_Selections_3").Cells(selectionRow, 1).Locked = True
        Range("Shots_Selections_4").Cells(selectionRow, 1).Locked = True
        Range("Shots_Selections_5").Cells(selectionRow, 1).Locked = True
        Range("Shots_Selections_6").Cells(selectionRow, 1).Locked = True
        Range("Shots_Combinations").Cells(selectionRow, 1).Locked = True
        Range("Shots_True_Prices").Cells(selectionRow, 1).Locked = True
        Range("Shots_Offer_Prices").Cells(selectionRow, 1).Locked = True
        Range("Shots_Selection_Names").Cells(selectionRow, 1).Locked = True
        
    Else
       
    End If
    
    selectionRow = selectionRow + 1
    
Loop

Application.ScreenUpdating = True
ThisWorkbook.Worksheets("Shots Selections").Protect

End Sub

Public Sub shotErrors()

Dim rowCount As Integer: rowCount = 1
Dim totalShots As String
Dim selectionName As String

Do Until Range("Shots_Combinations").Cells(rowCount, 1) = ""

    totalShots = Range("Shots_Combinations").Cells(rowCount, 1)
    selectionName = Range("Shots_Selection_Names").Cells(rowCount, 1)
    
    If StringMatch(selectionName, totalShots) = False Then
        MsgBox "Error with selection " & rowCount & " of shot combinations!"
        End
    Else
      rowCount = rowCount + 1
    End If
    
    Loop

End Sub



