Dim Message, Speak
Message=InputBox("Enter Text","Speak")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak Message