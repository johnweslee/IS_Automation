'Function Name : FindStringInFile()
'Description : To check if a given string in present at specfic line in a file
'Parameters : 
	'strFilePath - Path of the file
	'intLineNumber - Line number (Integer Value)
	'strStringToFind - String to find
'Return Value : 0 if true or -1 if false
'=================================================================================
Function FindStringInFileA(strFilePath, intLineNumber, strStringToFind)
	On Error Resume Next
	FindStringInFileA = -1
	Const For_Reading = 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFilePath, For_Reading)
	If objFSO.FileExists(strFilePath) Then
		For i = 1 To (intLineNumber - 1)
			objFile.SkipLine
		Next
		strLineContent = objFile.ReadLine
		If (InStr(strLineContent, strStringToFind) > 0) Then
			FindStringInFileA = 0
		End If
	End If
	Err.Clear
	Set objFSO = Nothing
	Set objFile = Nothing
End Function


Function FindStringInFileB(strFilePath, intLineNumber, strStringToFind)
	On Error Resume Next
	FindStringInFileB = -1
	Const For_Reading = 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFilePath, For_Reading)
	arrLines = Split(objFile.ReadAll, vbCrLf)
	If objFSO.FileExists(strFilePath) Then
		If UBound(arrLines) >= intLineNumber-1 Then
			If arrLines(intLineNumber-1) <> "" Then
				strLineContent = arrLines(intLineNumber-1)
			End If
		End If
		If (InStr(strLineContent, strStringToFind) > 0) Then
			FindStringInFileB = 0
		End If
	End If
	Err.Clear
	Set objFSO = Nothing
	Set objFile = Nothing
End Function
