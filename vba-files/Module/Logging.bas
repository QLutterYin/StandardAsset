Attribute VB_Name = "Logging"
Option Explicit


Sub LogInformation(LogMessage As String, Optional modulName As String)

    Dim LogFileName As String
    Dim FileNum As Integer
    Dim message As String
    
    LogFileName = ActiveWorkbook.path & "\log.log"
    message = "(" & Format(Now, "yyyy-mm-dd hh:mm::ss") & ")"
    If modulName = "" Then
        message = message & "[VBAproject]-INFO: "
    Else
        message = message & "[VBAproject::" & modulName & "]-INFO: "
    End If
    message = message & LogMessage
    
    message = Replace(message, "VBAproject", ActiveWorkbook.Name)
    
    FileNum = FreeFile ' next file number
    
    Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist
    
    Print #FileNum, message ' write information at the end of the text file
    
    Close #FileNum ' close the file

End Sub

