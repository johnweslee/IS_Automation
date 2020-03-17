'Code the auot renumber the steps in test script file

On Error Resume Next

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WsShell = CreateObject("WScript.Shell")

Const ForReading = 1, ForWriting = 2, TriStateFalse = 0
Dim strArray(), strFinal, InputFile, strLog, OutputFile

'CurrentWorkingDirectory
TestCaseFolder = Mid(WScript.ScriptFullName, 1, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))

'I/OFiles
InputFile = TestCaseFolder + "input.wsf.txt"
OutputFile = TestCaseFolder + "output.wsf.txt"

'Read contents from the Input File & store in a array variable
Set objTemp = FSO.OpenTextFile(InputFile, ForReading, False, TriStateFalse)
strLog = objTemp.ReadAll
objTemp.Close

intLength = Len(strLog) - 1
Redim strArray(intLength)
intStep = 1

For i = 0 to intLength
	strArray(i) = Mid(strLog, i + 1,1)
Next

'Renumber the steps
For i = 0 to intLength
	strChar = TypeName(CInt(strArray(i)))
	strChar1 = TypeName(CInt(strArray(i+1)))
	
	'for 2 Digit
	a = ((strChar = "Integer") AND (strChar1 = "Integer") AND (strArray(i-1) = " ") AND (strArray(i-2) = "p") AND (strArray(i-3) = "e") AND (strArray(i-4) = "t") AND (strArray(i-5)) = "S" AND (strArray(i-6) = chr(34)))
	'For 1 Digit
	b = ((strChar = "Integer") AND (strArray(i-1) = " ") AND (strArray(i-2) = "p") AND (strArray(i-3) = "e") AND (strArray(i-4) = "t") AND (strArray(i-5) = "S") AND (strArray(i-6) = chr(34)))
	
	If(b) Then
		strArray(i) = intStep
		strFinal = strFinal & strArray(i)
		If(a) Then
			strArray(i+1) = ""
			strFinal = strFinal & strArray(i+1)
		End If
		intStep = intStep + 1
		strChar = Empty
		strChar1 = Empty
	Else
		strFinal = strFinal & strArray(i)
		strChar = Empty
		strChar1 = Empty
	End If
Next

'Write the new contents to OutputFile
Set objTemp = FSO.OpenTextFile(OutputFile, ForWriting, True, TriStateFalse)
objTemp.Write strFinal
objTemp.Close

Set FSO = Nothing
Set WsShell = Nothing