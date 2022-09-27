Attribute VB_Name = "Test"

'namespace=xvba_unit_test/Test

'/*
'Sample file for put your test rotines here
'File on xvba_unit_test wil export o Excel/Access just on production
'*/
Public Sub index()
    dim x As String

    On Error GoTo ErrorHandle:

    'Children Function Call that will throw an Error and Raise to the Sub
    x = 1/0
    Xdebug.printx x

 ErrorHandle:
    Xdebug.errorSource = "pageConsoller.index"
    Xdebug.printError

End Sub