Attribute VB_Name = "ExportTrade"
Option Explicit
'Updated 2020-01-21
Sub ExportTrade() 'sheetName As String, waitOnReturn As Boolean, suppressOverwriteExistingWarning As Boolean)
    Dim sheetName As String: sheetName = Range("R2").Value
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim suppressOverwriteExistingWarning As Boolean: suppressOverwriteExistingWarning = True
    '''''''''''''''''''''''''''''''''''
    Dim tradeFolder As String
    Dim workingFolder As String
    Dim workingFileName As String
    Dim combineFileName As String
    Dim joinArray() As String
    Dim mergePathString As String
    Dim r As Long
    
    ' get paths
    tradeFolder = Application.ActiveWorkbook.Path & "\Report Exports\Trades\" & sheetName & "\"
    workingFolder = tradeFolder & "Working Files\"
    
    workingFileName = workingFolder & sheetName & "_" & Application.WorksheetFunction.Text(Sheets(sheetName).Range("R3").Value, "yyyy-mm-dd") & ".pdf"
    workingFileName = Replace(workingFileName, "/", "-")
    combineFileName = tradeFolder & sheetName & "_" & Application.WorksheetFunction.Text(Sheets(sheetName).Range("R3").Value, "yyyy-mm-dd") & ".pdf"
    combineFileName = Replace(combineFileName, "/", "-")
        
    'verify folders exists
    If Not DirExists(workingFolder) Then
        MyMkDir (workingFolder)
    End If
    
    ' if files exist kill them
    If FileExists(workingFileName) = True Then
        If suppressOverwriteExistingWarning = True Then
            Kill workingFileName
        Else
            r = MsgBox("Would you like to overwrite " & workingFileName & "?", vbYesNo + vbCritical, "Are you sure?")
            If r = vbYes Then
                Kill workingFileName
                r = 0
                Else
                r = 0
                Exit Sub
            End If
        End If
    End If
    If FileExists(combineFileName) = True Then
        If suppressOverwriteExistingWarning = True Then
            Kill combineFileName
        Else
            r = MsgBox("Would you like to overwrite " & combineFileName & "?", vbYesNo + vbCritical, "Are you sure?")
            If r = vbYes Then
                Kill combineFileName
                r = ""
                Else
                r = ""
                Exit Sub
            End If
        End If
    End If
    
    'export working file
    ThisWorkbook.Worksheets(sheetName).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        workingFileName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
        
    ' create merge array
    'ReDim joinArray(0)
    'joinArray(0) = workingFileName
    If Sheets(sheetName).Range("T2").Value <> "" Then
        mergePathString = workingFileName & "----" & Sheets(sheetName).Range("T2").Value
        joinArray = Split(mergePathString, "----")
        
    Else
        ReDim joinArray(0)
        joinArray(0) = workingFileName
    End If
    
    Call CombinePDFs(joinArray, combineFileName, waitOnReturn)
    
End Sub

Sub TestExportFormat()
    Dim nameFile As String: nameFile = Application.ActiveWorkbook.Path & "\Format Test.pdf"
    Dim sheetName As String: sheetName = "Cover"

    If IsFile(nameFile) = True Then
        'If MsgBox(nameFile & " already exists.  Replace it?" & vbNewLine & vbNewLine & "NOTE: Make sure the file is closed or you will get all kinds of errors.", vbYesNo) = vbNo Then Exit Sub
        Kill nameFile
    End If

    ThisWorkbook.Worksheets(sheetName).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        nameFile, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True

End Sub

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function
Public Sub MyMkDir(sPath As String)
    ' https://www.devhut.net/2011/09/15/vba-create-directory-structurecreate-multiple-directories/
    Dim iStart          As Integer
    Dim aDirs           As Variant
    Dim sCurDir         As String
    Dim i               As Integer
 
    If sPath <> "" Then
        aDirs = Split(sPath, "\")
        If Left(sPath, 2) = "\\" Then
            iStart = 3
        Else
            iStart = 1
        End If
 
        sCurDir = Left(sPath, InStr(iStart, sPath, "\"))
 
        For i = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(i) & "\"
            If Dir(sCurDir, vbDirectory) = vbNullString Then
                MkDir sCurDir
            End If
        Next i
    End If
End Sub
Function FileExists(FullFileName As String) As Boolean
' returns TRUE if the file exists
 FileExists = Len(Dir(FullFileName)) > 0
End Function
Sub CombinePDFs(inputPathArray As Variant, outputPath As String, waitOnReturn As Boolean)
    ' This is my trying to adapt the old combinepdfs from
    ' https://www.mrexcel.com/forum/excel-questions/870539-combine-pdfs-using-vba-2.html
    ' and
    ' https://stackoverflow.com/questions/15951837/wait-for-shell-command-to-complete
    ' and turn them into something that uses sejda-console because it is really fast
    ' and because some of the new pdf files versions weren't working with pdftk
    ' Written by Camron Walker 01/07/2019
    ' https://gitlab.com/camronwalker/uip-master
    
    'Dim inputPathArray
    'Dim outputPath As String: outputPath = "C:\Users\camron\code\uip-master\QuantityLinks\Output PDF File.pdf"
    'Dim waitOnReturn As Boolean: waitOnReturn = True
    'inputPathArray = Array("C:\Users\camron\code\uip-master\QuantityLinks\Test PDF File.pdf", "C:\Users\camron\code\uip-master\QuantityLinks\Test PDF File.pdf", "C:\Users\camron\code\uip-master\QuantityLinks\Test PDF File.pdf")
    
    Dim dejdaPath As String: dejdaPath = Application.ActiveWorkbook.Path & "\includes\sejda-console\bin\sejda-console.bat"
    Dim strShell As String
    Dim i As Long
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim windowStyle As Integer: windowStyle = 7
    
    strShell = """" & dejdaPath & """" & " merge -f"
    
    For i = LBound(inputPathArray) To UBound(inputPathArray)
        strShell = strShell & " """ & inputPathArray(i) & """"
    Next i
    
    strShell = strShell & " -o """ & outputPath & """ -a flatten --overwrite -b one_entry_each_doc > """ & Application.ActiveWorkbook.Path & "\debug.txt"""
    Debug.Print strShell
    
    wsh.Run strShell, windowStyle, waitOnReturn

    Set wsh = Nothing
End Sub
