VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalTable_Userform 
   Caption         =   "Items"
   ClientHeight    =   4344
   ClientLeft      =   135
   ClientTop       =   570
   ClientWidth     =   5475
   OleObjectBlob   =   "CalTable_Userform.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "CalTable_Userform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandOK_Click()
' This sub is used to calcute what-if data table manuelly.

' Author: Qi Lutter-Yin
' Date: 22.03.2021

    Dim startTimes
    Dim i As Integer
    Dim tblAddress As String, rowAddress As String, colAddress As String
    Dim startCells As Range, startCell
    Dim inputItems
    
    
    Call beforeMacro
'    LogInformation "Customized data table calculation starts.", "calTable_customized"
    
    delTable ' delete the existing tables so that it will not be updated if the file needs calculation.
    
    startTimes = Now
    
    
    With CalTable_Userform.ListBox1
        If .ListCount >= 1 Then
            For i = 0 To .ListCount - 1
'            Debug.Print .List(i)
                If .Selected(i) Then
'                    Debug.Print "selected: " & .List(i)
                    Application.StatusBar = "tables for " & .List(i) & " are being calculated."
                    Set startCells = getDTStarts(.List(i))  'item i selected
        '            startCells.Select
                    inputItems = Split(.List(i), ",")
                    Call IterateTables2(startCells, _
                                        Range(inputItems(0)).Address, _
                                        Range(inputItems(1)).Address)
                End If
            Next i
        End If
    End With
    
        
    Application.StatusBar = False
    
    'Debug.Print "Customized single thread takes: " & Now - startTimes
    
    ' log a message in this sub
'    LogInformation "Customized data table calculation finishes. It takes: " & _
    Format(Now - startTimes, "nn:ss"), "calTable_customized"
    CalTable_Userform.Hide
    
    Call afterMacro
End Sub

Private Sub CommandCancel_Click()
    CalTable_Userform.Hide
End Sub

Private Sub UserForm_Initialize()
    With ListBox1
        .AddItem "GDPflex,ERVflex"
        .AddItem "PP,GDPflex"
        .AddItem "PP,ERVflex"
        .AddItem "LTPPflex,Marginflex"
        .AddItem "PP,Multipleflex"
        .AddItem "PP,Quarterflex"
    End With
End Sub
