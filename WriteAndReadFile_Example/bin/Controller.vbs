dim ControllerIsInitialized
const INIT_DONE = 11


' ------------------------------------------------------------------------------------------------------------------------
'	loadController
'
'	Die Funktion initialisiert das Controller Skript einmalig.
'	Innerhalb der If Abfrage müssen alle externen Skripte eingetragen werden, welche durch den Controller 
'	verwendet werden sollen.
public sub loadController()
	if not ControllerIsInitialized = INIT_DONE then
		ControllerIsInitialized = INIT_DONE
		IncludeScript("bin/FA_FileAccessor.vbs")
	end if
end sub




' ------------------------------------------------------------------------------------------------------------------------
'	WriteTextToFile
'
'	@FilenameIBox		Text HTML-Input Element das den Dateinamen beinhaltet
'	@TextItext			Text HTML-Input Element das den Text beinhaltet, der in die Datei geschrieben werden soll
public sub WriteTextToFile()
	dim Filename , InputText
	Filename = document.getElementById("FileNameTextbox").value
	InputText = document.getElementById("FileTextTextbox").value
	
	call FA_WriteToFile(Filename,InputText)
end sub

' ------------------------------------------------------------------------------------------------------------------------
'	ShowFileContent
'
'	Zeigt den Inhalt der Datei an, deren Namen in der "FileNameTextbox" Input-Textbox eingegben ist.
public sub ShowFileContent()
	dim Filename
	Filename = document.getElementById("FileNameTextbox").value
	dim FileContent
	FileContent = FA_ReadFile(Filename)

	document.getElementById("FileContentOutput").innerText = FileContent
end sub



' ------------------------------------------------------------------------------------------------------------------------
'	IncludeScript
'		
'	Lädt externe VBS Datei in den Globalen Skriptkontext ein. 
'	Das Skript bindet somit externe .vbs Skripte ein.
private Sub IncludeScript(Byval filename)
  Dim codeToInclude
  Dim FileToInclude

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1
  Const OpenFileForReading = 1
  Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")

  'Check for existance of include
  If Not FSO.FileExists(filename) Then
    msgbox "Controller.vbs/IncludeScript/Include file not found."
    Set FSO = Nothing
    Exit Sub
  End If

  'open file to include
  Set FileToInclude = FSO.OpenTextFile(filename, ForReading, _
  FailIfNotExist, OpenAsDefault)

  'read all contet of the file
  codeToInclude = FileToInclude.ReadAll
  
  'close file after reading
  FileToInclude.Close
 
  'now cleanup the unused objects
  Set FSO = Nothing
  Set FileToInclude = Nothing

  'now execute code from include file
  ExecuteGlobal codeToInclude
End Sub