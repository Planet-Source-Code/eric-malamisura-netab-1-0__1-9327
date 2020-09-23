Attribute VB_Name = "modStuff"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function OpenIt(frm As Form, ToOpen As String)
On Error Resume Next
ShellExecute frm.hwnd, "Open", ToOpen, &O0, &O0, SW_NORMAL
End Function

Public Function CreateInternetShortcut(FileTitle As String, URL As String)
On Error Resume Next
Dim StrURLFile As String
Dim StrURLTarget As String
Dim FileNum As Integer
Dim Dialog As New cCommonDialog
Dialog.Filter = "URL Shortcuts (*.url)|*.url"
Dialog.FilterIndex = 1
Dialog.Filename = FileTitle
Dialog.FileTitle = FileTitle
Dialog.DialogTitle = "Save Shortcut To..."
Dialog.ShowSave
'initialise Variables
StrURLFile = FileTitle & ".url"
StrURLTarget = URL
FileNum = FreeFile
'Write the Internet Shortcut file
Open StrURLFile For Output As FileNum
Print #FileNum, "[InternetShortcut]"
Print #FileNum, "URL=" & StrURLTarget
Close FileNum
End Function

Public Function OpenURL()
On Error Resume Next
Dim Dialog As New cCommonDialog
Dialog.Filter = "Htm(*.htm)|*.htm|html(*.html)|*.html|ASP(*.asp)|*.asp"
Dialog.FilterIndex = 1
Dialog.DialogTitle = "Select The Web Document"
Dialog.ShowOpen
frmMain.IE(frmMain.Tabmain.SelectedTab).Navigate Dialog.Filename
End Function

