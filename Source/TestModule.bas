Attribute VB_Name = "TestModule"
Sub fortest()

TEMP = MsgBox("Do you reset the Grating and Slit position ?", vbYesNo + vbQuestion, "Confirmation")
If TEMP = vbYes Then
    Range("A2") = Range("A2").value & "Slit position : 0 mm" & vbCrLf
    Range("A2") = Range("A2").value & "Grating position : 0 nm" & vbCrLf
    Range("H4").value = 0
    Range("H8:H10").value = 0
    Call MOVE_WORKING_ABS_POSITION
    Call SlitSetPosition
End If

End Sub
