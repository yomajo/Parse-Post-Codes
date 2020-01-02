Attribute VB_Name = "AddCodesFromTxt"
Dim objShell As Object
Dim objExec As Object
Dim objOutput As Object
Dim oCmd As String
Dim sOutputStr As String
Dim sLine As String
Dim TxtFilePath As String
Dim PythonwExePath As String
Dim PyScriptPath As String

Sub AddPostalCodes()
'Add codes from picked txt file to Postal_Codes_Manager.xlsx via python script 'txt_to_excel.py'

Application.ScreenUpdating = False

'Getting path for txt file (passing it to Py as argument)
TxtFilePath = GetTXTFile

PythonwExePath = ThisWorkbook.Worksheets("PyPath").Range("A1").Value
PyScriptPath = ThisWorkbook.Path & "\txt_to_excel.py"

'Making sure path to pythonw.exe is provided
If InStr(1, ThisWorkbook.Worksheets("PyPath").Range("A1").Value, "Failed") > 0 Then
    MsgBox "Problem determining 'pythonw.exe' path on workbook open macro" & vbCrLf & vbCrLf & "Ending program", vbCritical
    End
End If

'Reset variable
sOutputStr = ""

'Cmd command to be executed:
oCmd = PythonwExePath & " " & """" & PyScriptPath & """" & " " & """" & TxtFilePath & """" & " "

'Creating and start shell:
Set objShell = VBA.CreateObject("Wscript.Shell")
Set objExec = objShell.Exec(oCmd)
Set objOutput = objExec.StdOut
'Reading output
While Not objOutput.AtEndOfStream
    sLine = objOutput.ReadLine
    If sLine <> "" Then sOutputStr = sOutputStr & sLine & vbCrLf
Wend
Debug.Print sOutputStr

'Resetting objects
Set objOutput = Nothing
Set objExec = Nothing
Set objShell = Nothing

If InStr(1, sOutputStr, "LOADED") > 0 Then
    MsgBox "Postal codes added 'Postal_Codes_Management.xlsx' file!", vbInformation
Else
    MsgBox "Python script failed to load postal codes. Check selected text file!", vbCritical
End If

End Sub
