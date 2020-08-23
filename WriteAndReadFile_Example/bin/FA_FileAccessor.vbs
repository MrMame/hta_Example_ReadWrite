
Const READ = 1, WRITE = 2, APPEND = 8
const CREATE_IF_MISSING = true, DONT_CREATE_IF_MISSING = false 

' ---------------------------------------------------------------------------------------------------
'	FA_WriteToFile
'
'	Schreibt den Inhalt in die Textdatei
'		@targetFilename		Zieldatei, in die der Text angehangen werden soll
'		@textToWrite		Text, der der Zieldatei angehangen werden soll
public sub FA_WriteToFile(targetFilename,textToWrite)
	dim filesys, filetxt
	Set filesys = CreateObject("Scripting.FileSystemObject")

	Set filetxt = filesys.OpenTextFile(targetFilename, APPEND, CREATE_IF_MISSING)
	filetxt.WriteLine(textToWrite)

	filetxt.Close	
	
	msgbox("Targetfilename=" & targetFilename & vbcrlf & _
			"textToWrite=" & textToWrite)

end sub

' ---------------------------------------------------------------------------------------------------
'	FA_ReadFile
'
'	Liest den Inhalt der Zieldatei aus und gibt diesen als Rückgabewert zurück
'		return		Inhalt der Zieldatei
public function FA_ReadFile(targetFilename)
	dim filesys, filetxt
	Set filesys = CreateObject("Scripting.FileSystemObject")

	on error resume next
	Set filetxt = filesys.OpenTextFile(targetFilename, READ, DONT_CREATE_IF_MISSING)
	
	if err.number > 0 then
		msgbox "Es konnte keine Datei mit dem namen '" & targetFilename & "' zum lesen gefunden werden!" & vbcrlf _
				& "Bitte pruefen Sie den Dateinamen."
		FA_ReadFile = "Fehler beim lesen der Datei"
	else
		FA_ReadFile = filetxt.ReadAll
		filetxt.Close	
	end if
end function