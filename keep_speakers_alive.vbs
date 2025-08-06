Option Explicit

' Based on: https://superuser.com/a/1214436
' Change the path and filename below 
Dim sound : sound = "C:\code\speakers\10Hz.wav"

Dim o : Set o = CreateObject("wmplayer.ocx")
With o
    .url = sound
    .controls.play
    While .playstate <> 1
        wscript.sleep 100
    Wend
    .close
End With

Set o = Nothing
