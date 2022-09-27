Attribute VB_Name = "MyRibbon"

'namespace=vba-files/ribbons

Public Sub ProcessRibbon(Control As IRibbonControl)
    Select Case Control.ID
        'call different macro based on button name pressed
        Case "g1_button1"
            Export.ExcelRangeToPowerPoint
        Case "g1_button2"
            Export2.ExcelRangeToPowerPointBank
        Case "g1_button3"
            Export2.ExcelRangeToPowerPointUpdateBank
        Case "g2_button1"
            One_Pager.OnePager
        Case "g2_button2"
            One_Pager.Onepager_POP
        Case "g3_button1"
            create_data_tables.calTable_main
        Case "g4_button1"
            extendInput1.extend_Input1
        Case "g4_button2"
            extendInput2.extend_input2
        Case "g4_button3"
            extendInput1.extend_Input1_plus
    End Select
End Sub
