Attribute VB_Name = "Misc"
Sub RefreshLinks()
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
End Sub
Sub OpenSettings()
    ActiveWorkbook.Worksheets("Settings").Activate
End Sub

