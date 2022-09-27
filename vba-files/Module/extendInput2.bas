Attribute VB_Name = "extendInput2"
Option Explicit

Const wk_names = "LU, RL, RN, IP, PCF, TDCF, AO, AL, IC"

Sub extend_input2()
' Description: The main sub used to extend input in the sheets wk_names.
    Dim noAssets As Integer, noRentrolls As Integer
    Dim refEntry As Range, refEntries As Range
    Dim firstMatch As String
    Dim wkName
    Dim what As String
    Dim NoRows As Integer
    Dim tmp
    
    Dim blocks() As clsBlock, block
    Dim noBlocks As Integer
    Dim i As Integer
    
    noBlocks = 0

    ' set environment to make the macro more quickly
    Call beforeMacro
      
    ' get number of Assets and number of rentrolls
    noRentrolls = Sheets("RR").Cells(Cells.Rows.Count, init_cells_RR.startCol).End(xlUp).Row - _
                  init_cells_RR.startRow + 1
    tmp = get_assets()
    noAssets = UBound(tmp) + 1
    
'    If noAssets < 2 Then
'        MsgBox "This macro cannot deal with a rentroll list with only one asset."
'        Exit Sub
'    End If
    

    
    ' All tabs:  initialize blocks
    For Each wkName In Split(wk_names, ", ")
        Sheets(wkName).Activate
        Application.StatusBar = wkName & " is in the process to extend rows."
        
        If wkName = "PCF" Or wkName = "TDCF" Then
            what = "Asset ID"
            NoRows = noAssets
        ElseIf wkName = "AL" Then
            what = "#"
            NoRows = noAssets
        ElseIf wkName = "AO" Then
            what = "Property ID"
            NoRows = noAssets
        ElseIf wkName = "IC" Then
            what = "ID"
            NoRows = noAssets
        Else
            what = "Unique Unit ID"
            NoRows = noRentrolls
        End If

        Set refEntries = FindAll(CStr(Cells.Address), what:=what, SearchDirection:=xlPrevious)
'        Debug.Print refEntries.Address
        For Each refEntry In refEntries
'            Debug.Print wkName & ":" & refEntry.Address
            ReDim Preserve blocks(noBlocks)
            Set blocks(noBlocks) = New clsBlock
            Set blocks(noBlocks).currentRng2(refEntry.Column) = refEntry.Offset(1, 0)
    '            Debug.Print blocks(noBlocks).Address
            ' change the number of rows in the blocks
            blocks(noBlocks).modifyRows NoRows
    '            Debug.Print blocks(noBlocks).Address
            noBlocks = noBlocks + 1
        Next
    
    Next
    
    ' change the formulas in the blocks
    For Each block In blocks
        Application.StatusBar = Split(block.currentRng.Parent.Name, "!")(0) & " is in the process to copy formulas."
        block.copyFormulas
    Next
    
    
    ' update FX und FW column till the end in the sheet AA. A new block should be added up.
    ReDim Preserve blocks(noBlocks)
    Sheets("AA").Activate
    Set blocks(noBlocks) = New clsBlock
    Set blocks(noBlocks).currentRng = get_AA_block()
    
    ' change the formulas in the blocks again
    For Each block In blocks
        Application.StatusBar = Split(block.currentRng.Parent.Name, "!")(0) & " is in the process to copy formulas for the second time."
        block.copyFormulas
    Next
    
    
    Application.StatusBar = False
    
    If MsgBox("Do you copy the original asset IDs into sheet AL? ", vbYesNo) = vbYes Then
        Call copyAssetIDs(Sheets("AL").Cells.Find("#").Offset(1, 0))
    End If
    
    If MsgBox("Do you copy the original asset IDs into sheet IC? ", vbYesNo) = vbYes Then
        Call copyAssetIDs(Sheets("IC").Cells.Find("ID Input").Offset(1, 0))
    End If

    Call turnOffBackup
    ' set environment back
    Sheets("RR").Activate
    Application.Calculate
    Call afterMacro
End Sub

Function get_AA_block() As Range
' Description: This is used to update formulas in AA again, since it is based on the sheet PCF
    Dim refEntry As Range, rightCell As Range, tmpRng As Range
    
    ' initial blocks for AA
    Sheets("AA").Activate
    
    Set refEntry = FindAll(Cells.Address, what:="Senior Loan").Offset(7, 0)
'    Set rightCell = Cells(refEntry.Row, 500).End(xlToLeft)
    Set rightCell = refEntry.Offset(0, 2)
    
    Set tmpRng = Range(refEntry, rightCell)
    Set get_AA_block = Application.Intersect(Range(tmpRng, tmpRng.End(xlDown)), _
                                             Range(Range("GDPSeniorMacro").EntireColumn, Range("AICSeniorMacro").EntireColumn))

End Function
