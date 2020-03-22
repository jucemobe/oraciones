'usage: cscript replace.vbs Filename "StringToFind" "stringToReplace"
 
Option Explicit
Dim fso,strFilename,strSearch,strReplace,objFile,oldContent,newContent, someString, objFS, objTS, strContents, iNumberOfLinesToDelete, iIndexToDeleteFrom, i,j, strPreces, strPrecess
Dim newContent1, newContent2, newContent3, newContent4, newContent5, newContent6, newContent7, newContent8, newContent9, newContent10
Dim newContent11, newContent12, newContent13, newContent14, newContent15, newContent16, newContent17, newContent18, newContent19,   xmlFichero
Dim semaforo, semaforo1, semaforo2, semaforo3, semaforo4, semaforo5, semaforo6
Dim strFichero
Dim tagsxml (11)
'tagsxml (0)="<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
tagsxml (0)="<?xml version=""1.0"" ?>"
'tagsxml (0)="<?xml version=" & chr(34) & "1.0" & chr(34) & " encoding=" & chr(34) & "ISO-8859-1" & chr(34) & "?>"
tagsxml (1)="<speak version=""1.1"" "
tagsxml (2)=" xmlns=""http://www.w3.org/2001/10/synthesis"""
tagsxml (3)=" xml:lang=""es-ES"">"
tagsxml (4)="<sentence>"
tagsxml (5)="<pron sym=""S e jn o r ""/> Senor</pron>"
tagsxml (6)="<silence msec=""200000""/>"
tagsxml (7)="<silence msec=""200000""/>"
tagsxml (8)="<silence msec=""200000""/>"
tagsxml (9)="<silence msec=""200000""/>"
tagsxml (10)="</sentence>"
tagsxml (11)="</speak>"
Dim precPerso (11)
precPerso (0)="Te pido Señor por la reconciliación con mi esposa y con mis hijos."
precPerso (1)="Gracias Señor por la salud de la nena, te pido que la sanes completamente"
precPerso (2)="Si la salud mental de Uriel soy yo que se sane completamente"
precPerso (3)="Gracias Señor porque Isa, Uriel y la nena están aquí en Madrid, vivos para toda la vida te pido señor que vuelvan a estar conmigo"
precPerso (4)="Gracias Señor por vender el coche"
precPerso (5)="Gracias Señor por el préstamo para pagar las tarjetas, y te pido señor planificar bien gastos."
precPerso (6)="Gracias Señor porque Isa se ha ocupado profesionalmente, Espíritu Santo ilumínala para que tome las decisiones adecuadas en su trabajo. Y en su vida"
precPerso (7)="Te pido señor por mi trabajo, que sea bueno "
precPerso (8)="Espíritu Santo ilumíname para tomar las decisiones adecuadas en mi trabajo y en mi vida"
precPerso (9)="Espíritu Santo ilumíname para entender lo que me dicen"
precPerso (10)="Espíritu Santo ilumíname para decir las palabras adecuadas y que expresen correctamente mis pensamientos lléname de elocuencia"
precPerso (11)="Si la clave es amarte a ti te amaré toda la vida, a través de la Virgen María."
Dim RegX
Set RegX = NEW RegExp
Dim MyString, SearchPattern, ReplacedText
'MyString = "Ocelots make good pets."
semaforo=true
semaforo1=true
semaforo2=true
semaforo3=true
semaforo4=true
semaforo5=true
semaforo6=true
SearchPattern = "</*[^>]*>"
'ReplaceString = ""
RegX.Pattern = SearchPattern
RegX.Global = True
'ReplacedText = RegX.Replace(MyString, ReplaceString)
'Response.Write(ReplacedText)
Dim fsocopia
Dim cadena_preces, cadena_padrenuestro, personal
'cadena_preces="Palabra hecha carne para que vivamos de ella"
'cadena_padrenuestro=" velando amorosamente por nosotros, nos atrevemos a decir"



strFilename=WScript.Arguments.Item(0)


cadena_preces=WScript.Arguments.Item(1)
cadena_padrenuestro=WScript.Arguments.Item(2)

personal=WScript.Arguments.Item(3)

'strSearch="<\/\*\[\^>\]\*>"
'strSearch="Hora"
strReplace=""
 
'Does file exist?
Set fso=CreateObject("Scripting.FileSystemObject")
if fso.FileExists(strFilename)=false then
   wscript.echo "file not found!"
   wscript.Quit
end if
 
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
 
'Write file
newContent1=RegX.replace(oldContent,strReplace)


SearchPattern = vBCr
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent2=RegX.replace(newContent1,"")


SearchPattern = vBLf
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=vBCrLf

'Write file
newContent3=RegX.replace(newContent2,vBCrLf)





set objFile=fso.OpenTextFile(strFilename,2)
objFile.Write newContent3
objFile.Close 



'set objFile=fso.OpenTextFile("c:\tmp\testigo\prba.txt",2)
'objFile.Write newContent3
'objFile.Close 




'string fileName = "file.txt"
someString = "Ir a la portada del sitio"

WScript.Echo "someString: " & someString

'
'string[] lines = File.ReadAllLines(strFilename);
'int found = -1;
'for (int i = 0; i < lines.Length; i++) {
'  if (lines[i].Contains(someString)) {
'    found = i;
'    break;
'  }
'}
'
' Delete First n Lines of a Text File


Dim arrLines
Const FOR_READING = 1
Const FOR_WRITING = 2
'strFileName = "C:\scripts\test.txt"
'int iNumberOfLinesToDelete = found

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

iNumberOfLinesToDelete=0
Do Until instr(arrLines(iNumberOfLinesToDelete),someString)
 'objTS.WriteLine arrLines(iNumberOfLinesToDelete)
 iNumberOfLinesToDelete=iNumberOfLinesToDelete+1
Loop

iNumberOfLinesToDelete=iNumberOfLinesToDelete+2

WScript.Echo "Number of lines to delete: " & iNumberOfLinesToDelete


For i=0 To UBound(arrLines)
 If i > (iNumberOfLinesToDelete - 1) Then
 objTS.WriteLine arrLines(i)
 End If
Next

objTS.Close

WScript.Echo "Finished 1 part" 

'string fileName = "file.txt";
someString = "volver al inicio"
'
'string[] lines = File.ReadAllLines(strFilename);
'int found = -1;
'for (int i = 0; i < lines.Length; i++) {
'  if (lines[i].Contains(someString)) {
'    found = i;
'    break;
'  }
'}


' Delete Last n Lines of a Text File



'Const FOR_READING = 1
'Const FOR_WRITING = 2
'strFileName = "C:\scripts\test.txt"
'iNumberOfLinesToDelete = lines.Length - found
WScript.Echo "creando objeto" 

Set objFS = CreateObject("Scripting.FileSystemObject")

WScript.Echo "abriendo fichero para leer" 

Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

WScript.Echo "cerrando fichero para leer" 


arrLines = Split(strContents, vbNewLine)

WScript.Echo "abriendo fichero para escribir" 


Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)
'iIndexToDeleteFrom = UBound(arrLines)- iNumberOfLinesToDelete + 1

WScript.Echo "calculando indice para comenzar a borrar" 


iIndexToDeleteFrom=0
Do Until instr(arrLines(iIndexToDeleteFrom),someString)
 'objTS.WriteLine arrLines(iIndexToDeleteFrom)
 iIndexToDeleteFrom=iIndexToDeleteFrom+1
Loop


WScript.Echo "Index to deletefrom: " & iIndexToDeleteFrom


For i=0 To UBound(arrLines)
 If i < iIndexToDeleteFrom Then
 objTS.WriteLine arrLines(i)
 End If
Next

WScript.Echo "borrado" 


objTS.Close


WScript.Echo "fichero cerrado" 



'Set fsocopia = CreateObject("Scripting.FileSystemObject")

'fsocopia.MoveFolder "C:\tmp\testigo\2020\03\21\laudes.txt", "C:\tmp\testigo\2020\03\21\laudes_001.txt"
'fsocopia.MoveFolder "C:\Users\rock\Desktop\TestFolder\BigEyeCat.jpg", "C:\Users\rock\Desktop\TestFolder2\FunnyCat.jpg"



WScript.Echo "comienzo de borrado varias cadenas" 


Set RegX = NEW RegExp

SearchPattern = "-se repite la antífona"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent1=RegX.replace(oldContent,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


SearchPattern = "Salmo 99: Alegría de los que entran en el templo"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent2=RegX.replace(newContent1,strReplace)


WScript.Echo "cadena borrada:" & SearchPattern

SearchPattern = "Salmo 23: Entrada solemne de Dios en su templo"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent3=RegX.replace(newContent2,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


SearchPattern = "Salmo 66: Que todos los pueblos alaben al Señor"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent4=RegX.replace(newContent3,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


SearchPattern = "\[Salmo 99\]"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent5=RegX.replace(newContent4,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


SearchPattern = "\[Salmo 23\]"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent6=RegX.replace(newContent5,strReplace)


WScript.Echo "cadena borrada:" & SearchPattern

SearchPattern = "\[Salmo 66\]"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent7=RegX.replace(newContent6,strReplace)


WScript.Echo "cadena borrada:" & SearchPattern

SearchPattern = "\[quitar\]"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent8=RegX.replace(newContent7,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


SearchPattern = " al inicio y al fin"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=" al inicio y al fin" & vBCrLf
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent9=RegX.replace(newContent8,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


SearchPattern = "Cántico \[en Español\] \[en Español\] \[en Latín\] \[en Latín\]"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent10=RegX.replace(newContent9,strReplace)

SearchPattern = "\[Salmo 94\]"
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
newContent11=RegX.replace(newContent10,strReplace)

WScript.Echo "cadena borrada:" & SearchPattern


'Write file
set objFile=fso.OpenTextFile(strFilename,2)
objFile.Write newContent11
objFile.Close 


WScript.Echo "fin de borrado varias cadenas" 


WScript.Echo "comienzo de preces" 


'string fileName = "file.txt";
'someString = "en quien el Padre ha querido recapitular todas las cosas"
'someString = "para hacer de nosotros criaturas nuevas"
someString = cadena_preces
'
'string[] lines = File.ReadAllLines(strFilename);
'int found = -1;
'for (int i = 0; i < lines.Length; i++) {
'  if (lines[i].Contains(someString)) {
'    found = i;
'    break;
'  }
'}


' Delete Last n Lines of a Text File



'Const FOR_READING = 1
'Const FOR_WRITING = 2
'strFileName = "C:\scripts\test.txt"
'iNumberOfLinesToDelete = lines.Length - found
WScript.Echo "creando objeto" 

Set objFS = CreateObject("Scripting.FileSystemObject")

WScript.Echo "abriendo fichero para leer" 

Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

WScript.Echo "cerrando fichero para leer" 


arrLines = Split(strContents, vbNewLine)

WScript.Echo "limite inferior lbound:" & lbound(arrLines)
WScript.Echo "limite superior ubound:" & ubound(arrLines)



WScript.Echo "abriendo fichero para escribir" 

WScript.Echo "vARIABLE someString:" & someString

Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)
'iIndexToDeleteFrom = UBound(arrLines)- iNumberOfLinesToDelete + 1

WScript.Echo "calculando indice para leer response" 


iIndexToDeleteFrom=0
Do Until instr(arrLines(iIndexToDeleteFrom),someString)
 'objTS.WriteLine arrLines(iIndexToDeleteFrom)
 iIndexToDeleteFrom=iIndexToDeleteFrom+1
Loop

WScript.Echo "Index menos 1 to response preces: " & iIndexToDeleteFrom




WScript.Echo "String preces_antes: " & arrLines(iIndexToDeleteFrom+1)




'strPreces=  aqString.Replace(arrLines(iIndexToDeleteFrom), ";", "-")  




'*******Set RegX = NEW RegExp
'*******
'*******SearchPattern = ":"
'*******'ReplaceString = ""
'*******RegX.Pattern = SearchPattern
'*******RegX.Global = True
'*******'ReplacedText = RegX.Replace(MyString, ReplaceString)
'*******'Response.Write(ReplacedText)
'*******
'*******
'*******
'*******
'*******
'*******'strFilename=WScript.Arguments.Item(0)
'*******'strSearch="<\/\*\[\^>\]\*>"
'*******'strSearch="Hora"
'*******strReplace="-"
'*******


'Set fso=CreateObject("Scripting.FileSystemObject")
'if fso.FileExists(strFilename)=false then
'   wscript.echo "file not found!"
'   wscript.Quit
'end if
 
'Read file
'set objFile=fso.OpenTextFile(strFilename,1)
'oldContent=objFile.ReadAll
 
'Write file
'strPreces=RegX.replace(arrLines(iIndexToDeleteFrom),strReplace)
strPreces=arrLines(iIndexToDeleteFrom+1)









WScript.Echo "String preces: " & strPreces


'*********strPrecess = Split(strPreces, Chr(45))
'*********
'*********For i=2 To UBound(strPrecess)
'*********WScript.Echo "String preces: " & i 
'*********WScript.Echo "String preces: " & strPrecess(i)
'*********WScript.Echo "String preces: " & strPrecess(1)
'*********Next







'WScript.Echo "String preces-10: " & arrLines(iIndexToDeleteFrom-10)
'WScript.Echo "String preces-9: " & arrLines(iIndexToDeleteFrom-9)
'WScript.Echo "String preces-8: " & arrLines(iIndexToDeleteFrom-8)
'WScript.Echo "String preces-7: " & arrLines(iIndexToDeleteFrom-7)
'WScript.Echo "String preces-6: " & arrLines(iIndexToDeleteFrom-6)
'WScript.Echo "String preces-5: " & arrLines(iIndexToDeleteFrom-5)
'WScript.Echo "String preces-4: " & arrLines(iIndexToDeleteFrom-4)
'WScript.Echo "String preces-3: " & arrLines(iIndexToDeleteFrom-3)
'WScript.Echo "String preces-2: " & arrLines(iIndexToDeleteFrom-2)
'WScript.Echo "String preces-1: " & arrLines(iIndexToDeleteFrom-1)
'WScript.Echo "String preces0: " & arrLines(iIndexToDeleteFrom)
'WScript.Echo "String preces1: " & arrLines(iIndexToDeleteFrom+1)
'WScript.Echo "String preces2: " & arrLines(iIndexToDeleteFrom+2)
'WScript.Echo "String preces3: " & arrLines(iIndexToDeleteFrom+3)
'WScript.Echo "String preces4: " & arrLines(iIndexToDeleteFrom+4)
'WScript.Echo "String preces5: " & arrLines(iIndexToDeleteFrom+5)
'WScript.Echo "String preces6: " & arrLines(iIndexToDeleteFrom+6)
'WScript.Echo "String preces7: " & arrLines(iIndexToDeleteFrom+7)
'WScript.Echo "String preces8: " & arrLines(iIndexToDeleteFrom+8)














WScript.Echo "comenzando proceso de preces " 


For i=0 To UBound(arrLines)
 If i < iIndexToDeleteFrom Then
    objTS.WriteLine arrLines(i)
 'else if instr(arrLines(i),"en quien el Padre ha querido recapitular todas las cosas") then
 'else if instr(arrLines(i),"para hacer de nosotros criaturas nuevas") then
 else if instr(arrLines(i),cadena_preces) then
            objTS.WriteLine arrLines(i)
'			objTS.WriteLine strPreces
      else if instr(arrLines(i),"-") then
                 objTS.WriteLine arrLines(i)
                 objTS.WriteLine strPreces
           else if instr(arrLines(i),"Se pueden añadir algunas intenciones libres.") then
		   if personal then
                    objTS.WriteLine precPerso (0)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (1)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (2)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (3)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (4)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (5)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (6)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (7)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (8)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (9)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (10)
                    objTS.WriteLine strPreces
                    objTS.WriteLine precPerso (11)
                    objTS.WriteLine strPreces
			end if
					semaforo=false
'                else if instr(arrLines(i),"Con el gozo que nos da el sabernos hijos de Dios, digamos con confianza") or semaforo  or instr(arrLines(i),"Digamos ahora, todos juntos, la ") then
'                else if instr(arrLines(i),"Con el gozo que nos da el sabernos hijos de Dios, digamos con confianza") or semaforo  or instr(arrLines(i),"Concluyamos nuestras ") then
                else if instr(arrLines(i),"Con el gozo que nos da el sabernos hijos de Dios, digamos con confianza") or semaforo  or instr(arrLines(i),cadena_padrenuestro) then
                         objTS.WriteLine arrLines(i)
                         semaforo=true
'     				else if instr(arrLines(i),"ver las intenciones de oración de ETF") then
'     					 else if instr(arrLines(i),"Por la Evangelización") then
'     						  else if instr(arrLines(i),"AMENAnonimo") then
'     							   else if instr(arrLines(i),"AMENAnónimo") then
'     								    else
'     										objTS.WriteLine arrLines(i)
'     									end if
'     							   end if
'     						  end if
'     					 end if
     				end if
     		  end if
           end if
       end if
 End If
Next

WScript.Echo "finalizado proceso de preces" 


objTS.Close


WScript.Echo "fichero cerrado" 


WScript.Echo "comienzo proceso acentos" 





strFilename=WScript.Arguments.Item(0)
 
'Does file exist?
Set fso=CreateObject("Scripting.FileSystemObject")
if fso.FileExists(strFilename)=false then
   wscript.echo "file not found!"
   wscript.Quit
end if

WScript.Echo "abriendo fichero para leer" 

 
'Read file
set objFile=fso.OpenTextFile(strFilename,1)
oldContent=objFile.ReadAll
objFile.Close

WScript.Echo "cerrando fichero para leer" 





'SearchPattern = chr(160)
SearchPattern = "á"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent1=RegX.replace(oldContent,strReplace)


SearchPattern = "é"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent2=RegX.replace(newContent1,strReplace)


SearchPattern = "í"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent3=RegX.replace(newContent2,strReplace)


SearchPattern = "ó"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent4=RegX.replace(newContent3,strReplace)


SearchPattern = "ú"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent5=RegX.replace(newContent4,strReplace)


SearchPattern = "Á"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent6=RegX.replace(newContent5,strReplace)


SearchPattern = "É"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent7=RegX.replace(newContent6,strReplace)


SearchPattern = "Í"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent8=RegX.replace(newContent7,strReplace)


SearchPattern = "Ó"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent9=RegX.replace(newContent8,strReplace)


SearchPattern = "Ú"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent10=RegX.replace(newContent9,strReplace)


SearchPattern = "ñ"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent11=RegX.replace(newContent10,strReplace)


SearchPattern = "Ñ"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace="�"

'Write file
newContent12=RegX.replace(newContent11,strReplace)



SearchPattern = ":\)"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent13=RegX.replace(newContent12,")")

SearchPattern = "&nbsp;"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent14=RegX.replace(newContent13,"")

SearchPattern = "†"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent15=RegX.replace(newContent14,"")


SearchPattern = "»"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent16=RegX.replace(newContent15,"")


SearchPattern = "«"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent17=RegX.replace(newContent16,"")


SearchPattern = "¿"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent18=RegX.replace(newContent17,"")


SearchPattern = "¡"
WScript.Echo SearchPattern
RegX.Pattern = SearchPattern
RegX.Global = True
strReplace=""

'Write file
newContent19=RegX.replace(newContent18,"")


'********SearchPattern = "†"
'********WScript.Echo SearchPattern
'********RegX.Pattern = SearchPattern
'********RegX.Global = True
'********strReplace=""
'********
'********'Write file
'********newContent15=RegX.replace(newContent14,"")


WScript.Echo "finalizado proceso acento" 



set objFile=fso.OpenTextFile(strFilename,2)
objFile.Write newContent19
objFile.Close 


WScript.Echo "fichero cerrado" 

WScript.Echo "comienzo de xml" 

WScript.Echo "creando objeto" 

Set objFS = CreateObject("Scripting.FileSystemObject")

WScript.Echo "abriendo fichero para leer" 

Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

WScript.Echo "cerrando fichero para leer" 


arrLines = Split(strContents, vbNewLine)

WScript.Echo "limite inferior lbound:" & lbound(arrLines)
WScript.Echo "limite superior ubound:" & ubound(arrLines)


strFichero = Split(strFilename, Chr(46))

WScript.Echo "abriendo fichero para escribir" 

WScript.Echo "limite inferior lbound:" & lbound(strFichero)
WScript.Echo "limite superior ubound:" & ubound(strFichero)


WScript.Echo "vARIABLE strFichero_name:" & strFichero(0)
WScript.Echo "vARIABLE strFichero_ext:" & strFichero(1)

xmlFichero=strFichero(0) & chr(46) & "xml"

WScript.Echo "vARIABLE fichero xml:" & xmlFichero

Set objTS = objFS.CreateTextFile(xmlFichero, True)
objTS.Close

Set objTS = objFS.OpenTextFile(xmlFichero, FOR_WRITING)

WScript.Echo "comenzando proceso de xml " 


For i=0 To UBound(arrLines)
 If i = 0 Then
    objTS.WriteLine tagsxml (0)
    objTS.WriteLine tagsxml (1)
    objTS.WriteLine tagsxml (2)
    objTS.WriteLine tagsxml (3)
    objTS.WriteLine tagsxml (4)
    'objTS.WriteLine tagsxml (5)
    objTS.WriteLine arrLines(i)
																																					WScript.Echo "pasamos por aqui 1i ---" & i
 else           if instr(arrLines(i),"Venid, aclamemos al Se") then
                    objTS.WriteLine tagsxml (6)
                    objTS.WriteLine arrLines(i)
																																					WScript.Echo "pasamos por aqui 2i ---" & i
                else             if instr(arrLines(i),"n en mi descanso.") then
                                     objTS.WriteLine arrLines(i)
                                     objTS.WriteLine arrLines(i+1)
                                     objTS.WriteLine arrLines(i+2)
                                     objTS.WriteLine tagsxml (6)
                                     semaforo1 = false
																																					WScript.Echo "pasamos por aqui 3i ---" & i
                                 else                 if instr(arrLines(i),"Ant:") or semaforo1 then
                                                                'objTS.WriteLine arrLines(i)
                                                                if semaforo1 = false then
                                                                      semaforo1=true
																	  semaforo4 = false
															      																						WScript.Echo "pasamos por aqui 4i ---" & i
                                                                end if
																if instr(arrLines(i),"or, Dios de Israel,") or semaforo4 then
                                                                    if semaforo4= false then
																	      objTS.WriteLine tagsxml (6)
																	'objTS.WriteLine arrLines(i)
																 						WScript.Echo "pasamos por aqui 5i ---" & i
                                                                          semaforo4=true
                                                                    semaforo5=false
																    end if
																    if instr(arrLines(i),"por el camino de la paz.") or semaforo5 then
																	                        if semaforo5=false then
                                                                                                'objTS.WriteLine arrLines(i)
                                                                                                'objTS.WriteLine tagsxml (6)
																							    semaforo5=true
                                                                                                semaforo2 = false
																								semaforo6=false
																							end if
                                                                                            'semaforo1 = false
																     						WScript.Echo "pasamos por aqui 6i ---" & i
                                                                                                          if instr(arrLines(i),"nus Deus Israel,") or semaforo2 then
                                                                                                                  'objTS.WriteLine arrLines(i)
                                                                                                                  'objTS.WriteLine tagsxml (6)
																    											  if semaforo2 = false then
																    											        semaforo2=true
																    													semaforo3=false
																    											  end if
																     						WScript.Echo "pasamos por aqui 7i ---" & i
                                                                                                                  if instr(arrLines(i),"um. Amen") or semaforo3 then
																    											         if semaforo3=false then
                                                                                                                               semaforo3=true
																														 else
																														       objTS.WriteLine arrLines(i)
																    													 end if
                                                                                                                         
                                                                                                                                          'objTS.WriteLine tagsxml (6)
																     						WScript.Echo "pasamos por aqui 8i ---" & i
                                                                                                                                                            'if instr(arrLines(i+1),"Demos gracias a Dios")  then
                                                                                                                                                            '     'objTS.WriteLine arrLines(i)
                                                                                                                                                            '     objTS.WriteLine tagsxml (10)
                                                                                                                                                            '     objTS.WriteLine tagsxml (11)
                                                                                                                                                            '      WScript.Echo "pasamos por aqui 9i ---" & i
                                                                                                                                                            'end if
                                                                                                                  end if
																                                          else 
																										       if semaforo6=false then
                                                                                                                    objTS.WriteLine arrLines(i)
                                                                                                                    objTS.WriteLine tagsxml (6)
																													semaforo6=true
																												else
																												     if instr(arrLines(i),"al de la cruz mientras se comienza a recitar") then
																												     else
                                                                                                                          objTS.WriteLine arrLines(i)
																													 end if
																												end if
                                                                                                          end if
																	else
																	    objTS.WriteLine arrLines(i)
                                                                    end if
																	
																	
																
																else                
                                                                objTS.WriteLine arrLines(i)
																end if
																
                                                      end if
                                 end if
                 end if
 End If
Next


objTS.WriteLine tagsxml (10)
objTS.WriteLine tagsxml (11)



WScript.Echo "finalizado proceso de xml" 


objTS.Close


WScript.Echo "fichero cerrado" 


WScript.Echo "fin de xml" 
