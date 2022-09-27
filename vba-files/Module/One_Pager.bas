Attribute VB_Name = "One_Pager"
Sub OnePager()
    '
    ' OnePager Macro
    '
    
    Call beforeMacro
    
    statusCal = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim MyCell As Range, MyRange As Range
    Dim wkCurrent As Workbook, wkOutput As Workbook
    
    Set wkCurrent = ActiveWorkbook
    Set wkOutput = Workbooks.Add
    
    wkCurrent.Activate
    Sheets("POP").Activate
    Set MyRange = Sheets("POP").Range("C75", Range("C1000").End(xlUp))
    
    For Each MyCell In MyRange
        'Calculate the values depending on Property ID
        wkCurrent.Sheets("POP").Range("E1").Value = MyCell.Value
        wkCurrent.Sheets("POP").Calculate
        wkCurrent.Sheets("ACF").Calculate
            
        'After the calculation for the whole workbook, the worksheets will be copied to the new workbook
        With wkOutput
'            Debug.Print MyCell.Value
            wkCurrent.Sheets(Array("POP", "ACF")).Copy after:=.Sheets(.Sheets.Count)
            
            .Sheets("POP").Name = MyCell & " - 1"
            .Sheets("ACF").Name = MyCell & " - 2"
        End With
    '    Logging.logINFO ("OnePager: " & MyCell.Value & " has been done!")
        
    Next MyCell
    wkOutput.Sheets(1).Delete
    wkOutput.BreakLink Name:=wkOutput.LinkSources(Type:=xlLinkTypeExcelLinks)(1), Type:=xlExcelLinks
    'Calculate
    Application.Calculation = statusCal
    
    Call afterMacro
End Sub

Sub Onepager_POP()
    '
    ' OnePager Macro
    ' This macro is almost identical to OnePager(), the only difference is that the worksheet("POP") is copied.
    '
    
    
    Call beforeMacro
    
    statusCal = Application.Calculation
    Application.Calculation = xlCalculationManual

    
    Dim MyCell As Range, MyRange As Range
    Dim wkCurrent As Workbook, wkOutput As Workbook
    
    Set wkCurrent = ActiveWorkbook
    Set wkOutput = Workbooks.Add
    
    wkCurrent.Activate
    Sheets("POP").Activate
    Set MyRange = Sheets("POP").Range("C75", Range("C1000").End(xlUp))
    
    For Each MyCell In MyRange
        'Calculate the values depending on Property ID
        wkCurrent.Sheets("POP").Range("E1").Value = MyCell.Value
        wkCurrent.Sheets("POP").Calculate
            
        'After the calculation for the whole workbook, the worksheets will be copied to the new workbook
        With wkOutput
    '        wkCurrent.Sheets("ACF").Copy After:=.Sheets(1)
    '        .Sheets("ACF").Name = MyCell & " - 1"
            
            wkCurrent.Sheets("POP").Copy after:=.Sheets(.Sheets.Count)
            .Sheets("POP").Name = MyCell
        End With
    '    Logging.logINFO ("OnePagerPOP: " & MyCell.Value & " has been done!")
    
        
    Next MyCell
    wkOutput.Sheets(1).Delete
    wkOutput.BreakLink Name:=wkOutput.LinkSources(Type:=xlLinkTypeExcelLinks)(1), Type:=xlExcelLinks
    'Calculate
    Application.Calculation = statusCal
    
    Call afterMacro

End Sub



