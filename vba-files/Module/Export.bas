Attribute VB_Name = "Export"
Option Explicit

Sub ExcelRangeToPowerPoint()
    
    Call beforeMacro
    
    Dim sourceRng As Range
    Dim ppt As Object
    Dim slide As Object
    Dim shp As Object, myshp As Object
    Dim targetShp As Object
    Dim shpName As String, pptName As String
    Dim i As Integer ' a variable for the loop
    Dim parentName As String, chartName As String
    Dim tmp, key
    Dim paragrahpNum As Integer
    
    Set ppt = get_PPT()
    
    If ppt Is Nothing Then Exit Sub
    
'    Application.Visible = False
        
    ' replace the graphics and text in the ppt
    For Each slide In ppt.slides
    
        ' replace the graphics in the ppt
        Call updateSlide(slide, createMappingGraphics())
        
        ' modify the text in the ppt
        Call updateSlideText(slide, createMappingText())
    Next
    
    ' modify the values of a table in ppt
    For Each shp In ppt.slides(3).Shapes
        If shp.Name = "Tbl KeyFacts" Then
            With shp.Table
                For i = 1 To 6
                    .cell(2, 1).Shape.TextFrame.TextRange.Paragraphs(i) = _
                    CStr(Sheets("OV").Range("M12").Offset(i - 1, 0).Value)
                Next
            End With
            Exit For
        End If
    Next
    
    'Clear The Clipboard
     Application.CutCopyMode = False
     
     ' tell the user that the export is finished.
     MsgBox "Export to PPT is finished."
'     Application.Visible = True
     
     Call afterMacro
End Sub


Private Function createMappingGraphics() As Object

' This dictionary defines the mapping between PPT pictures and their source charts in the excel.

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' adding items to the dictionary
    dict.Add key:="Sources", Item:="SU!B6:H9"
    dict.Add key:="Uses", Item:="SU!I6:K21"
    dict.Add key:="ES", Item:="ES!D7:O46"
    dict.Add key:="FinancingAssumptions", Item:="MA!J7:O18"
    dict.Add key:="PDStrategy", Item:="PD!BS81:CG86"
    dict.Add key:="CA", Item:="CA!B12:L28"
    dict.Add key:="PCA", Item:="PCA!B7:O24"
    dict.Add key:="CC", Item:="CC!B7:Q30"
    dict.Add key:="SensiGDP", Item:="ES!Q71:AF80"
    dict.Add key:="FFO", Item:="FO!D2:S37"
    dict.Add key:="TenantMix", Item:="TT!B7:J23"
    dict.Add key:="AssetList", Item:="AL!B6:Z111"

    
    ' THe items for charts should follow the syntax "worksheetname!Chart xxx"
    dict.Add key:="Grafik GEO", Item:="PD!Chart GEO"
    dict.Add key:="Grafik FMCG", Item:="PD!Chart FMCG"
    dict.Add key:="Grafik Strategy", Item:="PD!Chart Strategy"
    
    Set createMappingGraphics = dict

End Function

Private Function createMappingText() As Object

' This dictionary defines the mapping between PPT pictures and their source charts in the excel.

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' The key should be the names of certain shapes in the ppt. If the paragraphs in the shape should be replaced, the key should follow the syntax "shapename-num".
    dict.Add key:="Main Titel", Item:="Project " + Sheets("RR").Range("E7").Value ' The whole text in the shape named "Main Titel" should be replaced with new values.
    dict.Add key:="Portfolio Conformity Analysis incl. Reverse", Item:=Sheets("RR").Range("E7").Value + " " + "Portfolio Conformity Analysis incl. Reverse"
    dict.Add key:="Summary Text-2", Item:=Sheets("OV").Range("M8").Value ' The second paragraph in the shape named "Summary Text" should be replaced with new values.
    dict.Add key:="Summary Text-3", Item:=Sheets("OV").Range("M9").Value ' The third paragraph in the shape named "Summary Text" should be replaced with new values.
    dict.Add key:="S2 Date", Item:="*as per " & Sheets("GA").Range("G56").Value
    dict.Add key:="S3 Date", Item:="*as per " & Sheets("GA").Range("G56").Value

    Set createMappingText = dict

End Function


