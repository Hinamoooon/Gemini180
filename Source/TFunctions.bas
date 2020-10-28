Attribute VB_Name = "TFunctions"
'Tribial functions

Sub CLR_Progress()

Worksheets(1).Activate
Range("A2").value = ""
'Range("A2").MergeArea.ClearContents

End Sub

Sub UDG_Set()
UDG = Range("D7").value
End Sub
