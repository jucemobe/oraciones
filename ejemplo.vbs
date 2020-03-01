Option Explicit
Dim msg, dgeek, oFileStream, oVoice, i, text, filesys, newfolder


Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 'creates the wav file even if it is present in our folder
'Const SVSFDefault      = 0
'Const SVSFlagsAsync    = 1
'Const SVSFNLPSpeakPunc = 64
'Const SVSFIsFilename   = 4
'Const SVSFIsXML = 8


'Enum SpeechVoiceSpeakFlags
    'SpVoice Flags
Const    SVSFDefault = 0
Const    SVSFlagsAsync = 1
Const    SVSFPurgeBeforeSpeak = 2
Const    SVSFIsFilename = 4
Const    SVSFIsXML = 8
Const    SVSFIsNotXML = 16
Const    SVSFPersistXML = 32

    'Normalizer Flags
Const    SVSFNLPSpeakPunc = 64

    'TTS Format
'Const    SVSFParseSapi = 
'Const    SVSFParseSsml = 
'Const    SVSFParseAutoDetect = 

    'Masks
Const    SVSFNLPMask = 64
'Const    SVSFParseMask = 
Const    SVSFVoiceMask = 127
Const    SVSFUnusedFlags = -128
'End Enum






'usage: cscript replace.vbs Filename "StringToFind" "stringToReplace"
 
Dim fso,strFilename,strSearch,strReplace,objFile,oldContent,newContent, someString, objFS, objTS, strContents, iNumberOfLinesToDelete, iIndexToDeleteFrom
 
 
 
''*****Dim RegX
''*****Set RegX = NEW RegExp
''*****Dim MyString, SearchPattern, ReplacedText
''*****MyString = "Ocelots make good pets."
''*****'SearchPattern = "</*[^>]*>"
''*****SearchPattern = "[\x0A]+"
''*****'ReplaceString = "\x0D\x0A"
''*****RegX.Pattern = SearchPattern
''*****RegX.Global = True
''*****'ReplacedText = RegX.Replace(MyString, ReplaceString)
''*****'Response.Write(ReplacedText)
''*****
''*****
''*****
''*****
''*****
strFilename=WScript.Arguments.Item(0)

'Set oVoice = CreateObject("SAPI.SpVoice" )
'oVoice.rate = 2
'oVoice.Speak oldContent
'oVoice.Speak strFilename, SVSFIsFilename + SVSFlagsAsync
'oVoice.WaitUntilDone(10000)

Wscript.Echo ("file to read is:" & strFilename ) 'alerting user about the new folder

''*****'strSearch="\<\/\*\[\^\>\]\*\>"
''*****'strSearch="Hora"
''*****strReplace=""& vbCrLf
 
'Does file exist?
Set fso=CreateObject("Scripting.FileSystemObject")
if fso.FileExists(strFilename)=false then
   wscript.echo "file not found!"
   wscript.Quit
end if
 
'Read file
set objFile=fso.OpenTextFile(strFilename,1,0)
oldContent=objFile.ReadAll
 
'Write file
'*****newContent=RegX.replace(oldContent,strReplace)
'*****set objFile=fso.OpenTextFile(strFilename,2)
'*****objFile.Write newContent
objFile.Close 
'*****




'***Set dgeek=CreateObject("sapi.spvoice" )
'***i=hour(time)  'custom greeting
'***if i < 12 Then
'***i=("Good morning, I am Susy, Speech expert created by Daniel the geek" )
'***dgeek.Speak i
'***Else
'***i=("Good day, I am Susy, Speech expert created by Daniel the geek"  )
'***End If
'***
'***text=msgBox("Welcome - Dann v0.0.1 Text to audio converter" )
'***
'***msg=InputBox("Enter your text for conversion","Dann v0.0.1 Text to audio converter" )
'***
'***If msg = ("F***" ) Then 'word filtering add your preffered words
'***Err.Clear 
'***Wscript.Echo ("F words are not allowed, this response was trigerred because you entered an F word into the text field" ) 'display the rules
'***Else If msg = ("" ) Then 'setting a response if no text has been entered
'***dgeek.Speak ("You did not type anything for me to say, check back later, since your mind is blank" )
'***dgeek.WaitUntilDone(1000)
'***Else
'***dgeek.Speak msg
'***	End If
'***	End If

'creating a folder to export the sound file to

set filesys=CreateObject("Scripting.FileSystemObject" ) 
'checking if the folder does not exist 
If Not filesys.FolderExists("c:\tmp\testigo\" ) Then 
newfolder = filesys.CreateFolder("c:\tmp\testigo\" ) 'creating a custom folder

Wscript.Echo ("A new folder " & newfolder & " has been created" ) 'alerting user about the new folder
End If

'Saving the text entered as a wav

Set oFileStream = CreateObject("SAPI.SpFileStream" )

oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "C:\tmp\testigo\recording.wav", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice" )
Set oVoice.AudioOutputStream = oFileStream
oVoice.rate = 2
'oVoice.Speak oldContent
'oVoice.Speak oldContent, SVSFIsXML
'oVoice.Speak strFilename, SVSFIsFilename + SVSFIsXML + SVSFlagsAsync
'oVoice.Speak strFilename, SVSFIsFilename + SVSFlagsAsync
'oVoice.Speak strFilename, SVSFIsFilename
'oVoice.Speak strFilename, SVSFIsFilename + SVSFDefault
oVoice.Speak strFilename, SVSFIsFilename + SVSFIsXML 
oVoice.WaitUntilDone(10000)

oFileStream.Close