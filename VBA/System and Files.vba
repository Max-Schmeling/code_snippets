Function GetDesktop() As String
    ' Returns the path to the desktop
    Dim oWSHShell As Object
    On Error GoTo ErrorExit
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    Set oWSHShell = Nothing
    On Error GoTo 0
    Exit Function
    
ErrorExit:
    Set oWSHShell = Nothing
    On Error GoTo 0
    GetDesktop = ""
End Function


Function GetFolder(title As String) As String
    ' Opens the open folder dialog to let the user choose a folder
    ' Returns the selected folder path
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .title = title
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function


Function SaveWorkbookAs(Optional IntialFilename As String = "", Optional FileFormat As Integer = 51) As Boolean
    ' Opens the save workbook dialog to let the user save the workbook
    ' FileFormats: https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat
    ' - XlFileFormat.xlWorkbookDefault = 51
    ' - XlFileFormat.xlExcel12 = 50
    ' - XlFileFormat.xlOpenXMLWorkbookMacroEnabled = 52
    SaveWorkbookAs = Application.Dialogs(xlDialogSaveAs).Show(Arg1:=IntialFilename, Arg2:=XlFileFormat.xlWorkbookDefault)
End Function


Function FileExists(FilePath As String) As Boolean
    ' Returns True if <FilePath> exists. False if not or if an error is raised
    On Error Resume Next ' in case of illegal characters in <FilePath> or some system error
    FileExists = IIf(FilePath = vbNullString, False, Dir(FilePath, vbNormal) > "")
    On Error GoTo 0
End Function


Function DirectoryExists(DirectoryPath As String) As Boolean
    ' Returns True if the given <DirectoryPath> exists
    On Error Resume Next
    DirectoryExists = ((GetAttr(DirectoryPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function


Function IsOutlookOpen() As Boolean
    ' Returns True if Outlook is currently running
    Dim OLApp As Object
    On Error Resume Next
    Set OLApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If OLApp Is Nothing Then
        IsOutlookOpen = False
    Else
        IsOutlookOpen = True
    End If
End Function


Function IsOutlookInstalled(Optional ShowPrompt As Boolean = False) As Boolean
    ' Returns True if Outlook is installed, else False
    Dim OLApp As Object
    On Error GoTo NotInstalled
    Set OLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        Set xOLApp = Nothing
        IsOutlookInstalled = True
        Exit Function
    End If
NotInstalled:
    If ShowPrompt Then
        MsgBox "Outlook is not installed on this system.", vbExclamation
    End If
    IsOutlookInstalled = False
End Function


Sub CopyToClipboard(Text As String)
    ' Write <Text> to clipboard
    ' Gibberish on Windows 10 64bit

    On Error GoTo ErrorExit
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'MSForms DataObject
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrorExit:
    MsgBox "Could not copy '" & Text & "' to the clipboard.", vbOKOnly + vbCritical + vbDefaultButton1, "Error"
End Sub


Function GetClipboardText() As String
    ' Return string from clipboard
    ' Works on Windows 10 64bit

    'On Error GoTo ErrorExit
    Dim objData As Object
    Set objData = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'MSForms DataObject
    objData.GetFromClipboard
    GetClipboardText = objData.GetText()
    Exit Function
ErrorExit:
    GetClipboardText = ""
End Function
