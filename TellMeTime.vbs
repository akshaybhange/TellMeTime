Dim speaks, speech, hour

If hour(time) <= 12 Then
	If minute(time) = 30 Then
		speaks = "It's half past" & hour(time) 
	Else
		speaks = "It's " & hour(time) & " o'clock"
	End If
Else
	If minute(time) = 30 Then
		speaks = "It's half past" & hour(time) - 12
	Else
		speaks = "It's " & hour(time) - 12 & " o'clock"
	End If
End If

Set speech = CreateObject("sapi.spvoice")
speech.Speak speaks
