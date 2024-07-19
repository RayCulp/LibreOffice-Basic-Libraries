Option Explicit
'____________________________________________________________________________________________________________________________________________________________________________

Public Function ReadIniFileKey(sIniFile As String, sSection As String, sKey As String) As String
'******************************************************************************************
' Procedure Name: WriteIniFileKey
' Purpose: Returns the value of the specified key in an INI file
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sIniFile (String): The full Linux path to the INI file, i.e.
' Parameter sSection (String): The name of the section in which to look for the key
' Parameter sKey (String): The name of the key to look for
' Return value (String): The value of the key
' Usage example:
' Author: Ray Culp
' Date: 19.07.2024
' More information:
'******************************************************************************************
	' Declarations

		Dim sIniFileContent As String
		Dim sINILines() As String
		Dim sLine As String
		Dim i As Long
		Dim bSectionFound As Boolean
		' Dim bKeyExists As Boolean
		
	' Set up error handling
	
		On Error GoTo ErrorHandler 
		
	' Initializations
	
		sIniFileContent = ""
		bSectionFound = False
		' bKeyExists = False

	' Make sure the file exists
	
		If Not FileExists(sIniFile) Then
			ReadIniFileKey = "ERROR_INI_FILE_NOT_FOUND"
			Exit Function
		End If
		
	' Read the contents of the INI file

		sIniFileContent = ReadTextFile(sIniFile) ' Read the file into memory
		
	' Split the file contents into an array of individual lines
		
		sINILines() = Split(sIniFileContent, Chr(10))
		
	' Loop through all lines
		
		For i = LBound(sINILines) To UBound(sINILines)
		
		' Get the line
		
			sLine = sINILines(i)
		
		' Trim unnecessary spaces from the line
		
			sLine = Trim(sLine)
			
		' If we previously found the setcion we are looking for, but we have now encountered a new section,
		' then the key we are looking for in the section does not exist, so bail out.
			
			If bSectionFound And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
				Exit For
			End If
			
		' Check if this line contains the section we are looking for. If so, set bSectionFound = True
			
			If sLine = "[" & sSection & "]" Then
				bSectionFound = True
			End If
			
		' If bSectionFound = True and we haven't run into the start of a new section yet, then we are in
		' the section we want to be in, so test each line to see if it's the key we're looking for.
			
			If bSectionFound Then
			
			' Check whether the line is longer than the key we're looking for. If it isn't, then there is 
			' no need to test if the line contains the key.
			
				If Len(sLine) > Len(sKey) Then
				
				' Check whether the line starts with the key and "="
				
					If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
					
					' If it does, then this is the key we're looking for, so return everything to the right of the "="
					
						'bKeyExists = True
						
						ReadIniFileKey = Mid(sLine, InStr(sLine, "=") + 1)
						
						Exit For 
						
					End If
					
				End If
				
			End If
			
		Next i
		
CleanExit:

	Exit Function
		
ErrorHandler:

	Select Case True
	Case Err = 12345 ' A known/expected error has occurred
		' Do something here to handle this error
		Exit Function
	Case Else
		' Handle unknown/unexpected errors
		MsgBox "Error number: " & Err & Chr(13)  & Chr(13) & _
			"Error description: " & Error$ & Chr(13)  & Chr(13) & _
			"At line: " & Erl & Chr(13)  & Chr(13) & _
			"Date and time: " & Now , 48 ,"An unforeseen error occurred"
		Exit Function
	End Select
	
End Function
'____________________________________________________________________________________________________________________________________________________________________________

Public Function WriteIniFileKey(sIniFile As String, sSection As String, sKey As String, sValue As String) As Boolean
'******************************************************************************************
' Procedure Name: WriteIniFileKey
' Purpose: This function will:
'          -- Write information to an INI file
'          -- Create the INI file if it doesn't exist
'          -- Create the section and/or key if they don't exist
'          -- Update the value of an existing key
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sIniFile (String): The full linux path to the INI file, i.e. "/home/username/Documents/MyIniFile.ini"
' Parameter sSection (String): The section to search for and/or create the key in
' Parameter sKey (String): The key to create or update
' Parameter sValue (String): The value that should be assigned to the key
' Return value (Boolean): True if the operation succeeded, False if it failed
' Usage example:
' Author: Ray Culp
' Date: 19.07.2024
' More information:
'******************************************************************************************

	' Declarations

		Dim sIniFileContent As String
		Dim sINILines() As String
		Dim sLine As String
		Dim sNewLine As String
		Dim i As Long
		Dim bFileExists As Boolean
		Dim bInSection As Boolean
		Dim bKeyAdded As Boolean ' As opposed to section added or INI file created
		Dim bSectionFound As Boolean
		Dim bKeyExists As Boolean
		Dim lResult As Long 
		Dim lSectionStartLine As Long 
		Dim lSectionEndLine As Long 
		Dim lKeyLine As Long 
		
	' Enable error handling
	
		On Error GoTo ErrorHandler

	' Check whether the INI file already exists. If it does, the key will be updated or added.
	' If the INI file does not exist yet, it will be created using the section and key provided.
	
		If Not FileExists(sIniFile) Then
			sIniFileContent = ""
		Else
			sIniFileContent = ReadTextFile(sIniFile)
		End If
		
	' Split the INI file contents into an array of individual lines
	
		sINILines() = Split(sIniFileContent, Chr(10))
		
	' Clear the contents of sIniFileContent because it will be rebuilt using the lines in sINILines() and a new line if applicable
		
		sIniFileContent = ""
		
	' Loop through all lines and find the start and end of the section, if it exists
		
		For i = LBound(sINILines) To UBound(sINILines)
		
		' Get the contents of the current line
		
			sLine = sINILines(i)
		
		' Trim unnecessary spaces from the line
		
			sLine = Trim(sLine)

		' Test whether we already found the section header in a previous line
		
			If Not bSectionFound = True Then 
			
			' If the section hasn't been found yet, test whether this line contains the section header
			
				If sLine = "[" & sSection & "]" Then
				
				' If this line contains the section header we are looking for, then set bSectionFound to True so we don't check any further lines for the section
				
					bSectionFound = True
					
				' Save the starting line of the section
					
					lSectionStartLine = i
					
				End If 
			
			End If 
			
		' If the section header was already found and we are not still in the line that contains our header,
		' then test the to see if it is the next section header or the end of the file.
			
			If bSectionFound = True And i <> lSectionStartLine Then
			
			' Test whether this line contains a new section header

				If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
				
				' Save the end line of the section. This line contains a new section header, so subtract 1 to get the last line of the previous section

						lSectionEndLine = i - 1
						
				' We have all the information we need now, so exit the For Next loop

						Exit For
						
			' Otherwise, if we have reached the end of the file, this means the file only has one section

				ElseIf i = UBound(sINILines) Then 
				
				' Save the end line of the file, which is also the end of the section
				
					lSectionEndLine = i
					
				' We have all the information we need now, so exit the For Next loop
					
					Exit For 
					
				End If 
				
			End If
			
		Next i
		
	' If the section was found, check whether the key also exists
		
		If bSectionFound = True Then ' look for the key
		
			For i = lSectionStartLine + 1 To lSectionEndLine
			
				sLine = sINILines(i)
				
				If InStr(1, sLine, sKey, 1) > 0 Then 
					bKeyExists = True
					lKeyLine = i
					Exit For 
				End If 
			
			Next i
		
		End If 
		
	' We now know if the section and/or the key already exist. Depending on this information, we will
	' -- Create a new section and a new key-value pair 
	' -- Create a new key-value pair in an existing section
	' -- Update the value of an existing key-value pair in an existing section
	
	' Test whether the section was found
		
		If bSectionFound = True Then 
		
		' If the section was found, test whether the key was also found
		
			If bKeyExists = True Then
			
			' Both exist, so we will simply update the value of the key. 
			' Loop through the lines that contain the section.
				
				For i = LBound(sINILines) To UBound(sINILines)
				
				' Get the contents of the current line
				
					sLine = sINILines(i)
				
				' Trim unnecessary spaces from the line
				
					sLine = Trim(sLine)
					
				' Check whether this line is the line with the key we want to change
				
					If i = lKeyLine Then 
						sIniFileContent = sIniFileContent & sKey & "=" & sValue
					Else
						sIniFileContent = sIniFileContent & sLine 
					End If
					
					If i <> UBound(sINILines) Then 
						sIniFileContent = sIniFileContent & Chr(10) 
					End If 
					
				Next i 
				
			Else
			
			' The section  exists, but the key does not, so add the key at the end of the section
			
				For i = LBound(sINILines) To UBound(sINILines)
				
					' Get the contents of the current line
					
						sLine = sINILines(i)
					
					' Trim unnecessary spaces from the line
					
						sLine = Trim(sLine)
						
						If i <> lSectionEndLine Then 
							sIniFileContent = sIniFileContent & sLine & Chr(10)
						Else
							sIniFileContent = sIniFileContent & sLine & Chr(10)
							sIniFileContent = sIniFileContent & sKey & "=" & sValue & Chr(10)
						End If 
					
				Next i 
				
			End If 
			
		Else 

		' Neither the secion nor the key exists, so create both
		
			For i = LBound(sINILines) To UBound(sINILines)
			
			' Get the contents of the current line
			
				sLine = sINILines(i)
			
			' Trim unnecessary spaces from the line
			
				sLine = Trim(sLine)
			
				sIniFileContent = sIniFileContent  & sLine & Chr(10)
				
			Next i 

			sIniFileContent = sIniFileContent & "[" & sSection & "]" 
			sIniFileContent = sIniFileContent & Chr(10) & sKey & "=" & sValue
			
		End If 
		
	' Remove empty lines
		
		' Split the INI file contents into an array of individual lines
		
			sINILines() = Split(sIniFileContent, Chr(10))
			
		' Clear the contents of sIniFileContent because it will be rebuilt using the lines in sINILines() and a new line if applicable
			
			sIniFileContent = ""
	
		' Write sIniFileContent to the INI file
		
			For i = LBound(sINILines) To UBound(sINILines)
			
			' Get the contents of the current line
			
				sLine = sINILines(i)
			
			' Trim unnecessary spaces from the line
			
				sLine = Trim(sLine)
				
				If sLine <> "" Then 
				
					If i <> UBound(sINILines) Then 
						sIniFileContent = sIniFileContent & sLine & Chr(10) 
					Else
						sIniFileContent = sIniFileContent & sLine 
					End if
				
				End If 
				
			Next i 
			
	' Write sIniFileContent to the INI file
	
		lResult = WriteTextFile(sIniFileContent, sIniFile, True, "UTF-8")
		
	' Check return code
	
		If lResult = 0 Then 
			WriteIniFileKey = True
		ElseIf lResult = 1 Then
			MsgBox "The function WriteTextFile returned Error 1. This is usually due to an error in the file path. Please check." , 48 ,"File path error"
			WriteIniFileKey = False 
		End If 
	
CleanExit:

	Exit Function
		
ErrorHandler:

	Select Case True
	Case Err = 12345 ' A known/expected error has occurred
		' Do something here to handle this error
		Exit Function
	Case Else
		' Handle unknown/unexpected errors
		MsgBox "Error number: " & Err & Chr(13)  & Chr(13) & _
			"Error description: " & Error$ & Chr(13)  & Chr(13) & _
			"At line: " & Erl & Chr(13)  & Chr(13) & _
			"Date and time: " & Now , 48 ,"An unforeseen error occurred"
		Exit Function
	End Select
	
End Function
'____________________________________________________________________________________________________________________________________________________________________________

Sub Test_ReadIniFileKey
	
	Dim sIniFilePath As String

	Dim sIniSection As String 
	Dim sIniKey As String 
	Dim sValue As String 

	sIniFilePath = Environ("HOME") & "/Documents/MyIniFile.ini"
	sIniSection = "Section1"
	sIniKey = "Key1"
	
	sValue = ReadIniFileKey (sIniFilePath, sIniSection, sIniKey)
	
	MsgBox sValue 

End Sub
'____________________________________________________________________________________________________________________________________________________________________________

Sub Test_WriteIniFileKey
	
	Dim sIniFilePath As String
	Dim sIniSection As String 
	Dim sIniKey As String 
	Dim bResult As Boolean 

	sIniFilePath = Environ("HOME") & "/Documents/MyIniFile.ini"
	
	' Create a first section and key-value pair
	
		bResult = WriteIniFileKey (sIniFilePath, "Section1", "Key1", "Value for Section 1 Key 1")
		
	' Change the value of the key we just created
	
		bResult = WriteIniFileKey (sIniFilePath, "Section1", "Key1", "CHANGED Value for Section 1 Key 1")
		
	' Create a second key-value pair in the first section
	
		bResult = WriteIniFileKey (sIniFilePath, "Section1", "Key2", "Value for Section 1 Key 2")
		
	' Create a second section and a new key-value pair in it
	
		bResult = WriteIniFileKey (sIniFilePath, "Section2", "Key1", "Value for Section 2 Key 1")
		
	' Insert a new key at the end of the first section
	
		bResult = WriteIniFileKey (sIniFilePath, "Section1", "Key3", "Value for new key inserted at end of section 1")


End Sub
