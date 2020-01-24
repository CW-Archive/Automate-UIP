Attribute VB_Name = "OtherTrade"
Sub SelectTakeoffTrade()
    Dim fileToOpen As Variant
    Dim wbLocation As String
    Dim filesString As String
    
    wbLocation = ActiveWorkbook.Path
    
    ChDir wbLocation
    ChDrive wbLocation
    
    fileToOpen = Application.GetOpenFilename _
        (Title:="Select Quantity Links File", _
        FileFilter:="Document Files *.pdf (*.pdf),", MultiSelect:=True)
    
    ' add ---- between each file to store it in one cell as a string
    For Each f In fileToOpen
        filesString = filesString & "----" & f
    Next f
    
    filesString = Right(filesString, Len(filesString) - 4)
    
    'TODO:  Set the string to a range on the trade sheet
    Range("R2").Value = filesString
End Sub
