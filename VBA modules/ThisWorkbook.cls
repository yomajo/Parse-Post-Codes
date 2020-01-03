VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim objShell As Object
Dim fso As Object
Dim objExec As Object
Dim objOutput As Object
Dim oCmd As String
Dim sOutputStr As String
Dim sLine As String
Dim PythonwExePath As String

Private Sub Workbook_Open()
'Write Pythonw.exe PATH to A1 cell in PyPath worksheet on workbook open

Application.ScreenUpdating = False
'Reset variable
sOutputStr = ""

'Create new shell and filesystem objects
Set objShell = VBA.CreateObject("Wscript.Shell")
Set fso = VBA.CreateObject("Scripting.FileSystemObject")
'Cmd command to be executed:
oCmd = "Where Python"

'Start shell:
Set objExec = objShell.Exec(oCmd)
Set objOutput = objExec.StdOut
'Reading output
While Not objOutput.AtEndOfStream
    sLine = objOutput.ReadLine
    If sLine <> "" Then sOutputStr = sOutputStr & sLine
Wend

If InStr(1, sOutputStr, "python.exe") > 0 Then
    'using pythonw instead of python prevents cmd screen splash when executing py script
    PythonwExePath = Replace(sOutputStr, "python.exe", "pythonw.exe")
    'confirming file exists
    If Len(fso.GetFileName(PythonwExePath)) <> 0 Then
        ThisWorkbook.Worksheets("PyPath").Range("A1").Value = PythonwExePath
    Else
        ThisWorkbook.Worksheets("PyPath").Range("A1").Value = "Failed. Python installation does not have pythonw.exe file. Change in code on workbook open required."
    End If
Else
    ThisWorkbook.Worksheets("PyPath").Range("A1").Value = "Failed to get Python installation path"
End If

'Resetting objects
Set objOutput = Nothing
Set objExec = Nothing
Set objShell = Nothing

'Ensuring sheet is hidden
ThisWorkbook.Worksheets("PyPath").Visible = False

Application.ScreenUpdating = True
End Sub
