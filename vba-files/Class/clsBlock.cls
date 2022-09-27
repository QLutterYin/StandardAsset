VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' This class is used in the stardard asset model. It defines a special range.
' variables:
'   pCurrentRng: the defined range.
' Methods:
'   modifyRows: change the number of rows in the range for a given number.
'   copyFormulas: copy the formulas of the first rows to all other rows.

Private pCurrentRng As Range

Public Property Get currentRng()
    Set currentRng = pCurrentRng
End Property

Public Property Set currentRng(rng As Range)
    Set pCurrentRng = rng
End Property

Public Property Set currentRng2(locCol As Integer, firstEntry As Range)
' This property can be used if rows are selected as range

    Dim lastRow As Integer
    Dim tmpCell As Range

    lastRow = Sheets(firstEntry.Parent.Name).Cells(firstEntry.Row, locCol).End(xlDown).Row
    Set currentRng = Rows(firstEntry.Row & ":" & lastRow)

End Property

Sub modifyRows(NoRows As Integer)
' modify rows so that the nubmer of rows in the range is noRows
    Application.CutCopyMode = False
    

    Dim firstEntry  As Range
    Dim currentRows As Integer
    
    Sheets(currentRng.Parent.Name).Activate
    
    Set firstEntry = pCurrentRng.Resize(1)
    
    ' the cells in the middle should be removed.
    currentRows = pCurrentRng.Rows.Count
    If currentRows > 2 Then
        pCurrentRng.Resize(currentRows - 2).Offset(1, 0).Delete _
        shift:=xlUp
    End If
    
    ' insert new rows. After deleting the old cells, there are two rows left.
    If NoRows > 2 Then
        pCurrentRng.Resize(NoRows - 2).Offset(1, 0).Insert , _
        shift:=xlDown
    End If
    
    ' update pAddress
'    pAddress = "'" & pCurrentRng.Parent.Name & "'!" & _
'               pCurrentRng.Resize(NoRows).Address(External:=False)
    Set pCurrentRng = pCurrentRng.Resize(NoRows)
    Application.CutCopyMode = True
    
    ' scroll the window to cells(1,1)
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

End Sub

Sub copyFormulas()
' copy the formulas of the first entry to all other rows
    Application.CutCopyMode = False

    Dim firstEntry  As Range
    
    Sheets(currentRng.Parent.Name).Activate

    Set firstEntry = pCurrentRng.Resize(1)
    
    ' copy the formula from the first row to all other rows
    firstEntry.Copy
    pCurrentRng.PasteSpecial Paste:=xlPasteFormulas, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Application.CutCopyMode = True
    Cells(1, 1).Select
End Sub




