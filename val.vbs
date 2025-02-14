Set objShell = CreateObject("WScript.Shell")
Dim answer
Dim messages
Dim index

messages = Array("Please?", "Are you sure?","Pretty Please?",  "I will be very sad")
index = 0

Do
    answer = objShell.Popup("Will you be my valentine?", 0, "Message", 4 + 32)
    If answer = 6 Then
        Exit Do
    Else
        Do
            answer = objShell.Popup(messages(index), 0, "Message", 4 + 32)
            index = (index + 1) Mod 3
            If answer = 6 Then
                Exit Do
            End If
        Loop
    End If
Loop

Do
    answer = objShell.Popup("Yay! I'll see you tomo :)", 0, "Message", 0 + 32)
Loop While answer <> 1


