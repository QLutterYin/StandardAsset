Attribute VB_Name = "exportUtilities"
' Excel 2010 or later
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
    Option Explicit

Sub updateSlide(slide As Object, targetDictGraphics As Object)

    ' Description: This sub is used to update a slide based on a dictionary
    ' Params:
    ' slide: a slide in the presentation
    ' targetDictGraphcis: a dictionary used to map the shapes to ranges or charts in excel

    Dim shp As Object, targetShp As Object
    Dim shpName As String, parentName As String, chartName As String
    Dim tmp
    Dim sourceRng As Range

    For Each shp In slide.Shapes
        shpName = shp.Name
        If targetDictGraphics.exists(shpName) Then
            ' copy range to ppt
            tmp = Split(targetDictGraphics(shpName), "!")
            parentName = tmp(0)

            If Left(tmp(1), 5) = "Chart" Then
                chartName = tmp(1)
                Worksheets(parentName).ChartObjects(chartName).Activate
                ActiveChart.ChartArea.Copy
            ElseIf Left(tmp(1), 5) = "Group" Then
                chartName = tmp(1)
                Worksheets(parentName).Shapes.Range(Array(chartName)).Select
                Selection.Copy
            Else
                Set sourceRng = Range(targetDictGraphics(shpName))
                Worksheets(sourceRng.Parent.Name).Activate
                ActiveWindow.DisplayGridlines = False ' hide the gridlines
                sourceRng.Copy
            End If

            On Error GoTo errorHandler:
            slide.Shapes.PasteSpecial DataType:=2 '2 = ppPasteEnhancedMetafile

            Set targetShp = slide.Shapes(slide.Shapes.Count)

            ' set positions:
            With targetShp
                .Left = shp.Left
                .Top = shp.Top
                .Width = shp.Width
                If .Height > shp.Height Then '  And slide.slidenumber <> 2
                    .Height = shp.Height
                End If

                .ZOrder msoSendToBack
            End With

            ' delete the old one and rename the new one
            shp.Delete
            targetShp.Name = shpName
        End If
    Next shp

    Exit Sub

 errorHandler:
    ' if an error occurs for the pastespecial, this error handles three types of error

    ' 1. It might needs some time to copy. Let the macro wait 1 second.
    ' 2. It lost the copy. Solution: Copy again

    Debug.Print "Error Handling for the shape: " & shpName


    Sleep (1000) ' wait miliseconds to finish the copy

    If Left(tmp(1), 5) = "Chart" Then
        chartName = tmp(1)
        Worksheets(parentName).ChartObjects(chartName).Activate
        ActiveChart.ChartArea.Copy
    ElseIf Left(tmp(1), 5) = "Group" Then
        chartName = tmp(1)
        Worksheets(parentName).Shapes.Range(Array(chartName)).Select
        Selection.Copy
    Else
        Set sourceRng = Range(targetDictGraphics(shpName))
        Worksheets(sourceRng.Parent.Name).Activate
        ActiveWindow.DisplayGridlines = False ' hide the gridlines
        sourceRng.Copy
    End If

    slide.Shapes.PasteSpecial DataType:=2 '2 = ppPasteEnhancedMetafile

    Resume Next

End Sub

Sub updateSlideText(slide As Object, targetDictText As Object)

    ' Description: This sub is used to update a slide based on a dictionary
    ' Params:
    ' slide: a slide in the presentation
    ' targetDictText: a dictionary used to map the shapes to text in excel

    Dim paragrahpNum As Integer, i As Integer
    Dim key, tmp
    Dim shp As Object
    Dim shpName As String

    For i = slide.Shapes.Count To 1 Step -1

        Set shp = slide.Shapes(i)
        For Each key In targetDictText.keys
            tmp = Split(key, "-")
            If tmp(0) = shp.Name Then
                If UBound(tmp) = 1 Then
                    paragrahpNum = CInt(tmp(1))
                ElseIf UBound(tmp) = 0 Then
                    paragrahpNum = 1
                Else
                    MsgBox (shp.Name & ": The name of the text box is not allowed.")
                    GoTo nextIteration
                End If
                shp.TextFrame.TextRange.Paragraphs(paragrahpNum) = CStr(targetDictText(key))
            End If
 nextIteration:
        Next key

    Next i

End Sub


Sub insertPATAGraphics(slide As Object, paths)

    ' Description: This is used to update two pictures in the slide


    Dim shp As Object ' for the loop
    Dim i As Integer
    Dim targetShp(1) As Object

    On Error Resume Next
    For i = LBound(paths) To UBound(paths)

        Set targetShp(i) = slide.Shapes.AddPicture(paths(i), _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=60, Top:=35)

        If Err.Number = 0 Then
            For Each shp In slide.Shapes
                If shp.Name = "PATA" & CStr(i + 1) Then
                    With targetShp(i)
                        .Top = shp.Top
                        .Left = shp.Left
                        .Width = shp.Width
                        If .Height > shp.Height Then
                            .Height = shp.Height
                        End If
                    End With
                    shp.Name = "PATA_old"
                    shp.Delete
                End If

            Next shp
            targetShp(i).Name = "PATA" & CStr(i + 1)
        End If
    Next i

    On Error GoTo 0
End Sub
