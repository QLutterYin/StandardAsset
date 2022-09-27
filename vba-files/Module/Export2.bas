Attribute VB_Name = "Export2"
Option Explicit

' Excel 2010 or later
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub ExcelRangeToPowerPointBank()
    'PURPOSE: Copy/Paste An Excel Range Into a New PowerPoint Presentation
    'SOURCE: www.TheSpreadsheetGuru.com
    Dim ppt As Object, slide As Object
    Dim assetID
    Dim tmp
    Dim paths(1) As String
    Dim sFolder As String

    Call beforeMacro

    Calculate

    Set ppt = get_PPT()

    If ppt Is Nothing Then Exit Sub

        '    Application.Visible = False

        MsgBox "Please select the folder of the PATA pictures."
        sFolder = SelectFolder()

        For Each assetID In get_assets()
            If assetID <> 999 Then ' assetID = 999 is the backup-line, and it is excluded.

                ' For each asset id, generate and update the slide
                Set slide = ppt.slides(1).Duplicate
                slide.moveto ppt.slides.Count

                Call updateSlideBB(slide, assetID)

                ' update the PATA pictures
                ' the name of a picture should be assetID_1.jpg or assetID_2.jpg
                paths(0) = sFolder & "\" & assetID & "_1.jpg"
                paths(1) = sFolder & "\" & assetID & "_2.jpg"
                Call insertPATAGraphics(slide, paths)

            End If
            Next
            ppt.slides(1).Delete

            'Clear The Clipboard
            Application.CutCopyMode = False

            ' tell the user that the export is finished.
            MsgBox "Export to PPT is finished."
            '    Application.Visible = True

            Call afterMacro

End Sub

Sub ExcelRangeToPowerPointUpdateBank()
    'PURPOSE: Copy/Paste An Excel Range Into a New PowerPoint Presentation
    'SOURCE: www.TheSpreadsheetGuru.com
    Dim ppt As Object, slide As Object
    Dim assetID
    Dim tmp
    Dim paths(1) As String
    Dim sFolder As String
    Dim warningMsg As String

    Call beforeMacro

    Calculate

    Set ppt = get_PPT()

    If ppt Is Nothing Then Exit Sub

        '    Application.Visible = False

        MsgBox "Please select the folder of the PATA pictures."
        sFolder = SelectFolder()

        warningMsg = ""


        For Each slide In ppt.slides
            assetID = slide.Shapes.Title.TextFrame.TextRange.Text
            assetID = Split(Split(assetID, ", ")(0), ": ")(1)


            If IsInArray(assetID, get_assets()) = False Then
                ' check if the assetID is in the current model
                warningMsg = warningMsg & "Asset ID " & assetID & " does not exist." & vbNewLine
            Else
                Call updateSlideBB(slide, assetID)
                ' update the PATA pictures
                ' the name of a picture should be assetID_1.jpg or assetID_2.jpg
                paths(0) = sFolder & "\" & assetID & "_1.jpg"
                paths(1) = sFolder & "\" & assetID & "_2.jpg"
                Call insertPATAGraphics(slide, paths)
            End If

        Next slide

        If warningMsg <> "" Then MsgBox warningMsg

            'Clear The Clipboard
            Application.CutCopyMode = False

            ' tell the user that the export is finished.
            MsgBox "Export to PPT is finished."
            '    Application.Visible = True

            Call afterMacro

End Sub

Sub updateSlideBB(slide As Object, assetID)

    Dim shp As Object

    With Sheets("BB")
        .Range("C2") = assetID
        .Calculate
        slide.Shapes.Title.TextFrame.TextRange.Text = .Range("E2").Value2

        ' modify grafics
        Call updateSlide(slide, createMappingGraphics())

        ' modify date
        For Each shp In slide.Shapes
            If shp.Name = "date" Then
                shp.TextFrame.TextRange.Text = "(1) Calculated as of " & _
                Sheets("GA").Range("G56").Value

                Exit For
            End If
        Next shp

    End With

End Sub

Private Function createMappingGraphics() As Object

    ' This dictionary defines the mapping between PPT pictures and their source charts in the excel.

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' adding items to the dictionary
    dict.Add key:="AssetSummary", Item:="BB!M22:P40"
    dict.Add key:="TenancySchedule", Item:="BB!E6:K18"

    dict.Add key:="Grafik BB", Item:="BB!Group BB"

    Set createMappingGraphics = dict

End Function



