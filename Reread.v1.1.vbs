' NAME:	ReRead.vbs
' AUTHOR: Henrik Vestin Uppsala Biobank
' DATE: 2021
' VERSION: 1.1
' HISTORY: 1.0 initial version
'		   1.1 LIMS tycker inte om att läsa in modifierade plattfiler. Alla rader behöver finnas.
'			   Klockslag får inte diffa för mycket från tidigare inläst fil.
'		   
' COMMENT: Utgå från plattor som finns inlästa i LIMS och som behöver läsas in på nytt för att inkludera
'		   tomma rör.
'
'==================================================================

Option Explicit ' Force explicit variable declaration. 


Dim objFSO
Dim objRead
Dim objWrite
Dim fileNum

Dim strDate
Dim strTime
Dim timeStampReport
Dim modHours
Dim strHours

Dim FileLocation
Dim i
Dim y
Dim arrFiles
Dim arrFileValues
Dim strContents
Dim filePath
Dim strResultContents
Dim FileDestination
Dim newLine
Dim arrLines 
Dim arrLineValues
Dim tubeVolume
Dim Barcode
Dim lenfilepath

Set objFSo = CreateObject("Scripting.FileSystemObject")

'Function: open file and extract file path from each line.

FileLocation = "C:\temp\ReRead\"
FileLocation = FileLocation & "ST_plates_only.txt"
'FileLocation = FileLocation & "test.txt" 'dummypath remove later
	'wscript.Echo "Filelocation for platelist: " & FileLocation

Set objRead = objFSO.OpenTextFile(FileLocation, 1 , False)
strContents = objRead.ReadAll ' read file content 
objRead.Close
arrFiles = Split(strContents, vbCrLf) 

For i = 0 to UBound(arrFiles) ' 'for each line in textfile, open corresponding file
	arrFileValues = Split(arrFiles(i))
	filePath = arrFileValues(0)
	set objRead = objFSO.OpenTextFile(filePath, 1 , False)
	objRead.SkipLine ' skip first line in resultfile
	strResultContents = objRead.ReadAll
	objRead.Close
	arrLines = Split(strResultContents, vbCrLf)
	
	Barcode = Mid(filePath, 39, 10) ' grab barcode from filepath, grab 39th char and 10 chars after that. 
		
	FileDestination = "C:\temp\ReRead\Modified_files\" 'So that it stops appending previous barcode to filename.
	FileDestination = FileDestination & Barcode & ".csv"
	
	Set objWrite = objFSO.CreateTextFile(FileDestination, True)
	objWrite.WriteLine "containerbarcode;barcode;samplealias;containerposition;volume;concentration"
		For y = 0 to UBound(arrLines) - 5  
			arrLineValues = Split(arrLines(y), ";")
			newLine = arrLineValues(0) & ";" & arrLineValues(1) & ";" & arrLineValues(2) & ";" & arrLineValues(3) & ";" & arrLineValues(4) & ";" & arrLineValues(5)
			objWrite.WriteLine newLine
		Next
		if arrLineValues(0) = arr
		
		'v1.0 if tube has zero-volume it should be in the file, otherwise exclude tube.
		'For y = 0 to UBound(arrLines) - 5  
		'	arrLineValues = Split(arrLines(y), ";")
		'	newLine = arrLineValues(0) & ";" & arrLineValues(1) & ";" & arrLineValues(2) & ";" & arrLineValues(3) & ";" & arrLineValues(4) & ";" & arrLineValues(5)
		'	tubeVolume = arrLineValues(4)
		'	If tubeVolume = "0" Then
		'		objWrite.WriteLine newLine 
		'	End If
		'Next
	strDate = Mid(filePath, 50, 8)' grab 50th char and 8 following chars.
	strDate = Mid(StrDate, 1, 4) & "-" & Mid(strDate, 5)' inserting dash as separator for year-month
	strDate = Mid(StrDate, 1, 7) & "-" & Mid(strDate, 8)' inserting dash as separator for month-day
	strTime = Mid(filePath, 59, 6)' grab 59th char and 6 following chars.
	'strTime = Mid(strTime, 1, 2) & ":" & Mid(strTime, 3)' setting time from filename might confuse LIMS, since the time in the old resultfile might differ.
	strTime = "23:" & Mid(strTime, 3) ' setting a fixed time for hour to avoid confusing LIMS.
	strTime	= Mid(strTime, 1, 5) & ":" & Mid(strTime, 6)' inserting colon as time separator
	
	timeStampReport = strDate & " " & strTime

	objWrite.WriteLine "Last action tracked : " & timeStampReport
	objWrite.WriteLine "File created using : Tecan UBb"
	objWrite.Close
	
	fileNum = fileNum + 1
	wscript.Echo "File #" & fileNum ' i just count lines
Next	
	
