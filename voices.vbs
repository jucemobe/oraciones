Option Explicit
Dim Zira, David



Const SVSFDefault      = 0
Const SVSFlagsAsync    = 1
Const SVSFNLPSpeakPunc = 64
Const SVSFIsFilename   = 4


'Zira's Voice
Set Zira = CreateObject("SAPI.spVoice")
Set Zira.Voice = Zira.GetVoices("gender=female").Item(0)
'Zira.Rate = 2
Zira.Volume = 70

'David's Voice
Set David = CreateObject("SAPI.spVoice")
Set David.Voice = David.GetVoices("gender=female").Item(1)
'David.Rate = 2
David.Volume = 100

WScript.Echo "Zira voices count"

WScript.Echo Zira.GetVoices.Count

WScript.Echo "David voices count"

WScript.Echo David.GetVoices.Count

Zira.Speak "Mi nombre es Zira.", SVSFNLPSpeakPunc
Zira.Speak Zira.Voice.GetDescription
David.Speak "and My Name is David. It's nice to meet you!", SVSFNLPSpeakPunc
'David.Speak "Y mi nombre es David. Gusto en conocerte!", SVSFNLPSpeakPunc
David.Speak David.Voice.GetDescription
David.Speak "<pitch middle='-10'>What up My Robot?  "