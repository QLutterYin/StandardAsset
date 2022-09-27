Attribute VB_Name = "extendInput1"
Option Explicit
Const LAST_COL = 41 ' The last column in raw rentroll list file to be copied.
Const wk_names = "AA, TA" ' The tabs which need to be extended

Enum init_cells_RR ' The cell where the raw rentrolls should be copied to.
startRow = 23
startCol = 27
End Enum

Enum ref_cells_RR ' The formulas of the cells needs to be copied.
refRow = 23
startCol = 2
endCols = 26
End Enum

Sub extend_Input1()
    ' Description: The main sub used to extend input in the sheets RR, AA and TA
    Dim noAssets As Integer, noRentrolls As Integer ' number of assets and number of rentrolls
    Dim what As String, refEntry ' helpful variables to find refcells
    Dim wkName ' used to loop through the tabs
    Dim NoRows As Integer
    Dim tmp
    Dim blocks() As clsBlock, block
    Dim noBlocks As Integer
    Dim i As Integer
    Dim if_insert As Boolean
    Dim rentrollPath As String

    if_insert = False

        ReDim Preserve blocks(0)
        ' set environment to make the macro more quickly
        Call beforeMacro

        Sheets("RR").Activate

        ' get path of rentrolls and copy rentrolls into the target workbook
        rentrollPath = get_rentrolls_path()
        If rentrollPath = "" Then Exit Sub
            Call copyRentrolls(rentrollPath, if_insert)

            ' get number of Assets and number of rentrolls
            noRentrolls = Sheets("RR").Cells(Cells.Rows.Count, init_cells_RR.startCol).End(xlUp).Row - _
            init_cells_RR.startRow + 1
            tmp = get_assets()
            noAssets = UBound(tmp) + 1

            ' This is only for the format
            '    If noAssets < 2 Then
            '        MsgBox "This macro cannot deal with a rentroll list with only one asset."
            '        Exit Sub
            '    End If

            ' initial blocks for RR
            Application.StatusBar = "RR" & " is in the process to extend rows."
            Set refEntry = Range(Cells(ref_cells_RR.refRow, ref_cells_RR.startCol), _
            Cells(ref_cells_RR.refRow, ref_cells_RR.endCols))
            Set blocks(noBlocks) = New clsBlock
            Set blocks(noBlocks).currentRng = Range(refEntry, refEntry.Offset(10000, 0).End(xlUp))
            '    Debug.Print "Before modifyrows: " & blocks(noBlocks).currentRng.Address
            blocks(noBlocks).modifyRows noRentrolls
            '    Debug.Print "After modifyrows" & blocks(noBlocks).currentRng.Address

            ' initial blocks for AA and TA
            For Each wkName In Split(wk_names, ", ")
                Sheets(wkName).Activate
                Application.StatusBar = wkName & " is in the process to extend rows."
                noBlocks = noBlocks + 1
                ReDim Preserve blocks(noBlocks)

                If wkName = "AA" Then
                    what = "Asset ID" ' this is used to find the cell with the value "Asset ID"
                    NoRows = noAssets
                Else
                    what = "Unique Unit ID" ' this is only used for TA to find the cell with the value "Unique Unit ID"
                    NoRows = noRentrolls ' This is used both for TA and RA
                End If

                Set refEntry = FindAll(Cells.Address, what:="Asset ID")
                Set blocks(noBlocks) = New clsBlock
                Set blocks(noBlocks).currentRng2(refEntry.Column) = refEntry.Offset(2, 0)
                '            Debug.Print blocks(noBlocks).Address
                ' change the number of rows in the blocks
                blocks(noBlocks).modifyRows NoRows
                '            Debug.Print blocks(noBlocks).Address

                Next

                ' change the formulas in the blocks
                For Each block In blocks
                    Application.StatusBar = Split(block.currentRng.Parent.Name, "!")(0) & " is in the process to copy formulas."
                    block.copyFormulas
                    Next


                    Application.StatusBar = False

                    ' update the asset ID manuelly for "AA"
                    If MsgBox("Do you copy the original asset IDs into sheet AA? ", vbYesNo) = vbYes Then
                        Call copyAssetIDs(Sheets("AA").Cells.Find("Asset ID").Offset(2, 0))
                    End If

                    ' set environment back
                    Application.Calculate
                    Sheets("RR").Activate
                    Call afterMacro
End Sub

Sub extend_Input1_plus()
    ' Description: The main sub used to insert new entries in the sheets RR, AA and TA
    Dim noAssetsOld As Integer, noRentrollsOld As Integer ' number of assets and number of rentrolls
    Dim noAssetsNew As Integer, noRentrollsNew As Integer
    Dim what As String, refEntry ' helpful variables to find refcells
    Dim wkName ' used to loop through the tabs
    Dim NoRows As Integer
    Dim tmp
    Dim blocks() As clsBlock, block
    Dim noBlocks As Integer
    Dim i As Integer, lastRow As Integer
    Dim if_insert As Boolean
    Dim rrPath As String


    if_insert = True

        ReDim Preserve blocks(0)
        ' set environment to make the macro more quickly
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        Sheets("RR").Activate

        noRentrollsOld = Sheets("RR").Cells(Cells.Rows.Count, init_cells_RR.startCol).End(xlUp).Row - _
        init_cells_RR.startRow + 1
        tmp = get_assets()
        noAssetsOld = UBound(tmp) + 1

        rrPath = get_rentrolls_path()
        ' get path of rentrolls and copy rentrolls into the target workbook
        If rrPath = "" Then Exit Sub
            Call copyRentrolls(rrPath, if_insert)

            ' get number of Assets and number of rentrolls
            noRentrollsNew = Sheets("RR").Cells(Cells.Rows.Count, init_cells_RR.startCol).End(xlUp).Row - _
            init_cells_RR.startRow + 1
            tmp = get_assets()
            noAssetsNew = UBound(tmp) + 1

            ' initial blocks for RR
            Application.StatusBar = "RR" & " is in the process to insert rows."
            Set refEntry = Range(Cells(ref_cells_RR.refRow + noRentrollsOld - 2, ref_cells_RR.startCol), _
            Cells(ref_cells_RR.refRow + noRentrollsOld - 2, ref_cells_RR.endCols))
            Set blocks(noBlocks) = New clsBlock
            Set blocks(noBlocks).currentRng = Range(refEntry, refEntry.Offset(10000, 0).End(xlUp))
            '    Debug.Print "Before modifyrows: " & blocks(noBlocks).currentRng.Address
            blocks(noBlocks).modifyRows noRentrollsNew - noRentrollsOld + 2
            '    Debug.Print "After modifyrows" & blocks(noBlocks).currentRng.Address

            ' initial blocks for AA and TA
            For Each wkName In Split(wk_names, ", ")
                Sheets(wkName).Activate

                If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

                    Application.StatusBar = wkName & " is in the process to insert rows."
                    noBlocks = noBlocks + 1
                    ReDim Preserve blocks(noBlocks)

                    If wkName = "AA" Then
                        what = "Asset ID" ' this is used to find the cell with the value "Asset ID"
                        NoRows = noAssetsNew - noAssetsOld + 2
                    Else
                        what = "Unique Unit ID" ' this is only used for TA to find the cell with the value "Unique Unit ID"
                        NoRows = noRentrollsNew - noRentrollsOld + 2 ' This is used both for TA and RA
                    End If

                    Set refEntry = FindAll(Cells.Address, what:="Asset ID")
                    Set blocks(noBlocks) = New clsBlock
                    Set blocks(noBlocks).currentRng2(refEntry.Column) = refEntry.Offset(10000, 0).End(xlUp).Offset(-1, 0)
                    ' change the number of rows in the blocks
                    blocks(noBlocks).modifyRows NoRows
                    '            Debug.Print blocks(noBlocks).Address

                    Next

                    ' change the formulas in the blocks
                    For Each block In blocks
                        Application.StatusBar = Split(block.currentRng.Parent.Name, "!")(0) & " is in the process to copy formulas."
                        block.copyFormulas
                        '        Call DeleteContents(block.currentRng.Offset(1, 0))
                        Next


                        Application.StatusBar = False

                        ' update the asset ID manuelly for "AA"
                        If MsgBox("Do you copy the original asset IDs into sheet AA? ", vbYesNo) = vbYes Then
                            Call copyAssetIDs(Sheets("AA").Cells.Find("Asset ID").Offset(2, 0))
                        End If

                        ' set environment back
                        Application.Calculate
                        Sheets("RR").Activate
                        Application.ScreenUpdating = True
                        Application.DisplayAlerts = True

                        Call extend_input2
End Sub


Function get_rentrolls_path()
    ' Description: get the full path of rentroll list

    Dim fullpath As String
    If MsgBox("Chose the file of the rentrolls.", vbOKCancel) = vbCancel Then
        get_rentrolls_path = ""
    Else
        'Display a Dialog Box that allows to select a single file.
        'The path for the file picked will be stored in fullpath variable
        fullpath = getFile()

        Sheets("RR").Range("E12") = fullpath ' Show the path in the worksheet for the user.
        get_rentrolls_path = fullpath
    End If

End Function

Private Sub copyRentrolls(source As String, if_insert As Boolean)
    ' Description: copy rentroll list to sheet RR cell AA23
    ' Param:
    '   source: the full path of the rentroll list
    ' if_insert: a boolean. If True, the retrolls are add up, if false, the model will be generated completely new.

    Dim sourceWK As Workbook ' the workbook of the rentroll
    Dim targetWK As Workbook ' the current workbook
    Dim lastRow As Integer ' last rowin the source workbook
    Dim oldLastRow As Integer ' last row in the current workbook before rentrolls are inserted or added up.
    Dim targetCell As Range '
    Dim targetCellAdd As String


    Set targetWK = ActiveWorkbook
    Set sourceWK = Workbooks.Open(source)
    lastRow = Cells(Cells.Rows.Count, "C").End(xlUp).Row


    With targetWK
        Set targetCell = .Sheets("RR").Cells(init_cells_RR.startRow, _
        init_cells_RR.startCol)
        oldLastRow = targetCell.Offset(10000, 0).End(xlUp).Row

        If if_insert = True Then
            Set targetCell = .Sheets("RR").Cells(oldLastRow, init_cells_RR.startCol)
        End If

        ' copy the last row in the targetWK back to sourceWK so that there
        ' is also a backup entry in the new model. This is useful if new rentrolls
        ' are added up later.


        .Sheets("RR").Cells(oldLastRow, init_cells_RR.startCol) = "999" 'modify the asset ID of the backup
        .Sheets("RR").Range(.Sheets("RR").Cells(oldLastRow, init_cells_RR.startCol), _
        .Sheets("RR").Cells(oldLastRow, Cells.Columns.Count)).Copy _
        sourceWK.Worksheets(1).Cells(lastRow + 1, 3)

        targetCellAdd = targetCell.Address

        ' delete the old data
        If if_insert = False Then
            .Sheets("RR").Range(targetCell.Offset(1, 0), _
            .Sheets("RR").Cells(oldLastRow, Cells.Columns.Count)) _
            .Delete
        End If

    End With

    ' insert the new data
    sourceWK.Worksheets(1).Range(Cells(4, 3), Cells(lastRow + 1, LAST_COL)).Copy _
    targetWK.Sheets("RR").Range(targetCellAdd)

    ' close source workbook
    sourceWK.Close SaveChanges:=False
End Sub


