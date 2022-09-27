Attribute VB_Name = "create_data_tables"
Option Explicit

Enum dtSize ' The size of a sensi table
    nrow = 7
    ncol = 7
End Enum


Function getDTStarts(Optional ByVal dtName As String) As Range
' Description: This function finds all start cells to calculate sensi table
' Params:
'   dtName: If given, only the corresponding start cells will be returned.

    Dim resultRng As Range, startCells As Range
    Dim sName, dtNames
    
    dtNames = CalTable_Userform.ListBox1.List
    
    With Sheets("GA")
        
        Set resultRng = Nothing
        For Each sName In dtNames
            
            Select Case sName
                Case "GDPflex,ERVflex"
                    Set startCells = .Range("H731,R731,AB731")
                Case "PP,GDPflex"
                    Set startCells = .Range("H742,R742,AB742") '.Range("H741", "R741", "AB741")
                Case "PP,ERVflex"
                    Set startCells = .Range("H753,R753,AB753") ' .Range("H752", "R752", "AB752")
                Case "LTPPflex,Marginflex"
                    Set startCells = .Range("R763,AB763")  ' .Range("R762", "AB762")
                Case "PP,Multipleflex"
                    Set startCells = .Range("H774,R774,AB774") '.Range("H773", "R773", "AB773")
                Case "PP,Quarterflex"
                    Set startCells = .Range("H786,R786,AB786") '.Range("H785", "R785", "AB785")
                Case Else
            End Select
            
            If (dtName = sName) Then
               Set resultRng = startCells
               Exit For
            End If
            
            Set resultRng = union(resultRng, startCells)
        Next
    
    End With

    Set getDTStarts = resultRng

End Function

Sub calTable_main()
' This sub is used to calcute what-if data table manuelly.

' Author: Qi Lutter-Yin
' Date: 22.03.2021

Dim startTimes
Dim i As Integer
Dim tblAddress As String, rowAddress As String, colAddress As String
Dim startCells As Range, startCell
Dim inputItems


'LogInformation "Customized data table calculation starts.", "calTable_customized"

delTable ' delete the existing tables so that it will not be updated if the file needs calculation.

Sheets("GA").Activate

CalTable_Userform.Show
End Sub

Public Sub delTable()
    Dim startCells As Range, startCell
    
    Set startCells = getDTStarts
    
    For Each startCell In startCells
        startCell.Offset(1, 1).Resize(dtSize.nrow, dtSize.ncol).ClearContents
    Next startCell
    
End Sub

Sub IterateTables2(startCells As Range, rowAddress, colAddress)
'
' Proof-of-Concept code
' for faster calculation of a 2-D What-If data table
'

' Params:
'   startCells: a range of multiple cells
    Dim rngRowCell As Range
    Dim rngColCell As Range
    Dim startCell
    Dim varRowSet As Variant
    Dim varColSet As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim lCalcMode As Long
    Dim i As Long, j As Long, k As Long
    Dim varStartRowVal As Variant
    Dim varStartColVal As Variant
    Dim dTime As Double
    
    Set rngRowCell = Range(rowAddress)
    Set rngColCell = Range(colAddress)
    
    nRows = dtSize.nrow          ' number of rows in the Column of what-if values
    nCols = dtSize.ncol       ' number of columns in the row of what-if values
    
    lCalcMode = Application.Calculation ' Set environment
    Application.Calculation = xlCalculationManual


    ' get row and column arrays of what-if values
    '
    varRowSet = startCells(1, 1).Offset(0, 1).Resize(, nCols).Value2
'    startCells(1, 1).Offset(0, 1).Resize(, nCols).Select
    varColSet = startCells(1, 1).Offset(1, 0).Resize(nRows).Value2

    '
    ' initial start values
    '
    varStartRowVal = rngRowCell.Value2
    varStartColVal = rngColCell.Value2
    '
    ' calculate result for each what-if values pair
    '
    For j = 1 To nRows
        For k = 1 To nCols
                '
                ' set values for this iteration, recalc, store result
                '
                rngRowCell.Value2 = varRowSet(1, k)
                rngColCell.Value2 = varColSet(j, 1)
                Application.Calculate
                
                For Each startCell In startCells
                    startCell.Offset(j, k) = startCell.Value2
                Next startCell
                
        Next k
    Next j

    '
    ' reset back to initial values & recalc
    '
    rngRowCell.Value2 = varStartRowVal
    rngColCell.Value2 = varStartColVal
    Application.Calculation = lCalcMode
    Application.Calculate

        
End Sub
