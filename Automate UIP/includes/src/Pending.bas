Attribute VB_Name = "Pending"
Sub FeaturePending(buttonName As String)
    Dim e
    e = MsgBox("The " & buttonName & " button is not yet available.  Please update Automate UIP and try again.", vbExclamation, "Error - Feature Not Yet Available")
End Sub
