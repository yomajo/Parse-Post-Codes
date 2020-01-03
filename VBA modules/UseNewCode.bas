Attribute VB_Name = "UseNewCode"
Dim WorkingWb As Workbook
Dim PostalCodesWb As Workbook
Dim MainWs As Worksheet
Dim FreeCodesWs As Worksheet
Dim ExpiredCodesWs As Worksheet

Dim PythonwExePath As String
Dim PyScript As String
Dim PostalCodesFilePath As String
Dim WorkingWbName As String

Dim LastExpiredCodesCol As Long
Dim LastExpiredCodesRow As Long

Sub ProgramThatUsesNewCode()
'Insert new post code from 'Postal_Codes_Manager.xlsx', move used code to another sheet, edit code so its scannable by barcode scanner

Application.ScreenUpdating = False

PostalCodesFilePath = ThisWorkbook.Path & "\Postal_Codes_Manager.xlsx"

If Len(Dir(PostalCodesFilePath)) = 0 Then
    'Postal_Codes_Manager.xlsx does not exist, warn, terminate
    MsgBox "Postal_Codes_Manager.xlsx" & " Is missing or was moved. Program terminated.", vbExclamation, "File not Found"
    End
End If

Set WorkingWb = ThisWorkbook
Set MainWs = WorkingWb.Worksheets("Main")
Set PostalCodesWb = Workbooks.Open(PostalCodesFilePath, Local:=True)
Set FreeCodesWs = PostalCodesWb.Worksheets("Free Codes")
Set ExpiredCodesWs = PostalCodesWb.Worksheets("Expired Codes")
FreeCodes = FreeCodesWs.Cells(1048576, 1).End(xlUp).Row
WorkingWb.Activate

'Handling case of empty free codes sheet:
If FreeCodes = 1 And FreeCodesWs.Cells(FreeCodes, 1).Value = "" Then
    FreeCodes = 0
End If

'Checking to see if there are enough postal codes in PostalCodesFileName + parent warning check at 5000 remaining codes
If FreeCodes < 100 Then
    MsgBox "Less than 100 Postal Codes left in " & Dir(PostalCodesFilePath) & vbCrLf & "Press 'OK' to Continue", vbInformation, "Postal Codes running out"
    If FreeCodes = 0 Then
        answer = MsgBox("Postal_Code_Manager.xlsx does not have any postal codes" & vbCrLf & "Do you want to load new codes now?", vbYesNo + vbExclamation, "No Postal Codes Available")
        If answer = vbYes Then
            'close workbook and add postal codes with python
            PostalCodesWb.Close savechanges:=False
            Call AddCodesFromTxt.AddPostalCodes
            WorkingWb.Activate
            Call ProgramThatUsesNewCode
        Else:
            'User cancelled addition of new postal codes
            End
        End If
    End If
End If

'Copy Postal code
FreeCodesWs.Range("A1").Copy
'Paste Postal code and make it readable by barcode scanners adding starting and ending asterisks
MainWs.Range("C5").PasteSpecial xlPasteValues
MainWs.Range("C8").Value = MainWs.Range("C5").Value
MainWs.Range("C8").Value = "*" & MainWs.Range("C8").Value & "*"

'Determining where to move used postal code
LastExpiredCodesCol = ExpiredCodesWs.UsedRange.Columns(ExpiredCodesWs.UsedRange.Columns.Count).Column
LastExpiredCodesRow = ExpiredCodesWs.Cells(1048576, LastExpiredCodesCol).End(xlUp).Row
'Start new column and at row=1 when last row reaches desired number
If LastExpiredCodesRow = 5000 Then
    LastExpiredCodesCol = LastExpiredCodesCol + 1
    LastExpiredCodesRow = 0
End If

'Move used postal code to from 'Free Codes' to 'Expired Codes' Sheet
FreeCodesWs.Range("A1").Cut Destination:=ExpiredCodesWs.Cells(LastExpiredCodesRow + 1, LastExpiredCodesCol)
'Delete empty row
FreeCodesWs.rows("1:1").Delete shift:=xlUp

'Saving and closing PostalCodesFileName, shifting focus
PostalCodesWb.Close savechanges:=True
MainWs.Activate

Application.ScreenUpdating = True

MsgBox "New Code added"

End Sub

