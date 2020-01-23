Attribute VB_Name = "LinkFile"
Sub OpenQuantityLinkFile()
    Dim filePath As String: filePath = Range("Quantity_Link_File_Path").Value
    If filePath = "" Then Exit Sub
    If FileExists = False Then Exit Sub
    Workbook.Open filePath
End Sub

Sub SelectQuantityLinkFile()
    Dim strFileToOpen As String
    Dim wbLocation As String
    
    wbLocation = ActiveWorkbook.Path
    
    ChDir wbLocation
    ChDrive wbLocation
    
    strFileToOpen = Application.GetOpenFilename _
        (Title:="Select Quantity Links File", _
        FileFilter:="Excel Files *.xls* (*.xls*),")

    Range("Quantity_Link_File_Path").Value = strFileToOpen
End Sub
