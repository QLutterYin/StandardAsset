Attribute VB_Name = "utilities"
Option Explicit
'Uses Range.Find to get a range of all find results within a worksheet
' Same as Find All from search dialog box. It starts from the last finding.
'
Function FindAll(rngAddress As String, what As Variant, Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False) As Range
    Dim SearchResult As Range
    Dim firstMatch As String
    Dim rng As Range

    Set rng = Range(rngAddress)
    With rng
        Set SearchResult = .Find(what, , LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
        If Not SearchResult Is Nothing Then
            firstMatch = SearchResult.Address
            Do
                If FindAll Is Nothing Then
                    Set FindAll = SearchResult
                Else
                    Set FindAll = union(FindAll, SearchResult)
                End If
                Set SearchResult = .FindPrevious(SearchResult)
            Loop While Not SearchResult Is Nothing And SearchResult.Address <> firstMatch
        End If
    End With
End Function

' Get unique values for given array
Function unique(inputArr)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(inputArr) To UBound(inputArr)
        d(inputArr(i)) = 1
    Next i
    unique = d.keys

End Function

' get the unique values of assets
Function get_assets()

    Dim startCell As Range, endCell As Range
    Set startCell = Sheets("RR").Cells(init_cells_RR.startRow, init_cells_RR.startCol)
    Set endCell = startCell.Offset(10000, 0).End(xlUp)

    If startCell.Address = endCell.Address Then
        get_assets = Array(startCell.Value2)
    Else
        get_assets = unique(Application.Transpose(Range(startCell, endCell)))
    End If

End Function
Sub copyAssetIDs(targetCell As Range)
    Dim assetArr
    Dim i As Integer

    assetArr = get_assets()

    For i = LBound(assetArr) To UBound(assetArr)
        targetCell.Offset(i, 0) = assetArr(i)
    Next i

End Sub
' union ranges
Function union(ParamArray rgs() As Variant) As Range
    Dim i As Long
    For i = 0 To UBound(rgs())
        If Not rgs(i) Is Nothing Then
            If union Is Nothing Then Set union = rgs(i) Else Set union = Application.union(union, rgs(i))
            End If
        Next i
End Function

Sub DeleteContents(rng As Range)
    ' This sub is used to delete the contents in some columns. It is not used at the moment.

    '
    '    Application.FindFormat.Clear
    '    Application.FindFormat.Interior.Color = 16772300
    '    rng.Replace "", "", SearchFormat:=True
    '    Application.FindFormat.Clear

    Dim keepCols
    Dim deleteRng As Range, col As Range

    Application.FindFormat.Clear
    Application.FindFormat.Interior.Color = 16772300

    Select Case rng.Parent.Name
     Case "TA"
        keepCols = Array("G", "BB")
     Case "AA"
        keepCols = Array("E", "AB", "BH", "BK", "BN", "BQ", "FZ", "GK")
     Case Else
        Set deleteRng = rng
    End Select

    If Not IsEmpty(keepCols) Then
        For Each col In rng.Columns
            ' if the column is not in the keepCols array, it will be deleted later
            If IsError(Application.Match(CStr(Split(col.Address, "$")(1)), keepCols, 0)) Then
                Set deleteRng = union(deleteRng, Application.Intersect(rng, col))
                '                deleteRng.Select
            End If
        Next col
    End If

    deleteRng.Replace "", "", SearchFormat:=True
    Application.FindFormat.Clear

End Sub

Sub turnOffBackup()

    Dim targetCell As Range

    With Worksheets("TA")
        Set targetCell = .Range("C1000").End(xlUp)
        targetCell.Offset(0, 4) = "OFF"
    End With

    With Worksheets("AA")
        Set targetCell = .Range("C1000").End(xlUp)
        targetCell.Offset(0, 2) = "OFF"
    End With

End Sub


Sub beforeMacro()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Sub afterMacro()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Function get_PPT()

    Dim PowerPointApp As Object
    Dim pptName As String

    If MsgBox("Choose the template of the PPT. Please make sure that the presentation has proper names for the graphics.", vbOKCancel) = vbOK Then

        'Create an Instance of PowerPoint
        On Error Resume Next

        'Is PowerPoint already opened?
        Set PowerPointApp = GetObject(Class:="PowerPoint.Application")

        'Clear the error between errors
        Err.Clear

        'If PowerPoint is not already open then open PowerPoint
        If PowerPointApp Is Nothing Then Set PowerPointApp = CreateObject(Class:="PowerPoint.Application")

            'Handle if the PowerPoint Application is not found
            If Err.Number = 429 Then
                MsgBox "PowerPoint could not be found, aborting."
                Exit Function
            End If

            On Error GoTo 0

            'open a presentation
            pptName = getFile

            Set get_PPT = PowerPointApp.Presentations.Open(pptName)
        Else
            Set get_PPT = Nothing
        End If

End Function

Function getFile()
    Dim fileName As String
    With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .Filters.Add "All files", "*.*"
        'Show the dialog box
        .Show
        'Store in fullpath variable
        fileName = .SelectedItems.Item(1)
    End With

    getFile = fileName
End Function

Function SelectFolder()
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With

    SelectFolder = sFolder

End Function


Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Or CStr(element) = CStr(valToBeFound) Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
 IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function
