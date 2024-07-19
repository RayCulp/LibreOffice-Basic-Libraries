Option Explicit

'____________________________________________________________________________________________________________________________________________________________________________

Public Function ReadTextFile(ByVal sFilePath As String) As String 
'******************************************************************************************
' Procedure Name: ReadTextFile
' Purpose: Read and return the contents of a text file
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePath (String): The full path to the text file to read
' Return value (String): The contents of the text file
' Usage example: sFileContents = ReadTextFile("/home/username/Documents/MyText.txt")
' Author: Ray Culp
' Date: 26.06.2024
' More information:
'******************************************************************************************

	' Declarations
	
		Dim oSimpleFileAccess As Object
		Dim oTextInputStream As Object
		Dim oInputStream As Object   
		Dim sBuffer As string
		Dim lDelimiters() As Long
		Dim sFileURL As String 
	
	' Set up error handling

		On Error GoTo ErrorHandler
	
	' Create Simple File Access service
	
		oSimpleFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	
	' Convert file path to URL
		
		sFileURL = ConvertToUrl(sFilePath)
		
	' Creat Text Input Stream service
	
		oTextInputStream = CreateUnoService("com.sun.star.io.TextInputStream")
	
	' Open the file for reading using Simple File Access. Returns an InputStream
	' See https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ucb_1_1XSimpleFileAccess.html#a3cc4816f38cb20913837d1735b2b2d6d
	
		oInputStream = oSimpleFileAccess.OpenFileRead(sFileURL)
		
	' "Plug" the InputStream object into the TextInputStream (at least this is how I interpret what is happening).
	' See https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1io_1_1XActiveDataSink.html#ab2f54d6003382d17d74ee7f748bbd3ba
		
		oTextInputStream.SetInputStream(oInputStream)

	' Read the contents of the file into the string buffer
	
		sBuffer = oTextInputStream.readString(lDelimiters(), False)
	
	' Close the Text Input Stream
	
		oTextInputStream.CloseInput()
	
	' Return the contents of the file
	
		ReadTextFile = sBuffer
	
CleanExit:
		
	Exit Function
		
ErrorHandler:

	Select Case True
	Case Err = 12345 ' A known/expected error has occurred
		' Do something
		Exit Function
	Case Else
		' Handle unknown/unexpected errors
		MsgBox "Error number: " & Err & Chr(13)  & Chr(13) & _
		"Error description: " & Error$ & Chr(13)  & Chr(13) & _
		"At line: " & Erl & Chr(13)  & Chr(13) & _
		"Date and time: " & Now , 48 ,"An unforeseen error occurred"
	End Select
	
End Function 

'____________________________________________________________________________________________________________________________________________________________________________

Public Function WriteTextFile(ByVal sTextToWrite As String, ByVal sFilePath As String, ByVal bOverwrite As Boolean, Optional ByVal sEncoding As String) As Long
'******************************************************************************************
' Procedure Name: WriteTextFile
' Purpose: Write text to a file
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sTextToWrite (String): The text to write to the file
' Parameter sFilePath (String): The full path to the file
' Parameter iOverwrite (Boolean): Determines whether the function overwrites an existing file: 
'		True = Overwrite
'		False = Do not overwrite
' Parameter sEncoding (String): The name of the character set to use for encoding
'		See https://www.iana.org/assignments/character-sets/character-sets.xhtml for a full list of
'		character sets. If no character set is provided, function defaults to "UTF-8". 
'		See end of this module for a short list of common characters sets.
' Return values: 
'		0 = No error
'		1 = Type: com.sun.star.ucb.InteractiveAugmentedIOException
'			   Message: a folder could not be created at ./ucbhelper/source/provider/cancelcommandexecution.cxx:84."
'		2 = File already exists, but bOverwrite was set to False
' Usage example: 
' Author: Ray Culp
' Date: 26.06.2024
' More information: 
'		See also: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1io_1_1XTextOutputStream.html
'******************************************************************************************

	' Declarations

		Dim oSimpleFileAccess As Object
		Dim oOutputStream As Object
		Dim oTextOutputStream As Object
		Dim sError As String 
		
	' Set up error handling
		
		On Error GoTo ErrorHandler 
		
	' Create a SimpleFileAccess service
	
		oSimpleFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
		
    ' Check if the file exists
    
	    If oSimpleFileAccess.exists(sFilePath) Then
	    
	    ' Check whether we should overwrite it
	    
	    	If bOverwrite = true Then 
	    
	        	oSimpleFileAccess.kill(sFilePath)
	        	
	        Else 
	        	
	        	' If not, return an error and bail out
	        	WriteTextFile = 2
	        	Exit Function 
	        	
	        End If 
	        
	    End If
		
	' Open the file for writing
		
		oOutputStream = oSimpleFileAccess.openFileWrite(sFilePath)
		
	' If not encoding was provided, default to UTF-8
	
		If IsMissing(sEncoding) Then 
	
			sEncoding = "UTF-8"
		
		End If 
		
	' Create a TextOutputStream with UTF-8 character set
	
		oTextOutputStream = createUnoService("com.sun.star.io.TextOutputStream")
		
		oTextOutputStream.setOutputStream(oOutputStream)
		
		oTextOutputStream.setEncoding(sEncoding)
		
	' Write the text to the file
		
		oTextOutputStream.writeString(sTextToWrite)
		
	' Close the TextOutputStream
	
		oTextOutputStream.closeOutput()
		
	' Return 0
	
		WriteTextFile = 0
		
CleanExit:
		
	Exit Function
		
ErrorHandler:

	Select Case True
	Case Err = 12345 ' A known/expected error has occurred
		' Do something
		Exit Function
	Case Err = 1
		sError = Error$
		If InStr (1, Error$, "folder could not be created", 1) > 0 Then 
			' There is most likely something wrong with the file path
			WriteTextFile = 1
			Exit Function
		End If 
	Case Else
		' Handle unknown/unexpected errors
		MsgBox "Error number: " & Err & Chr(13)  & Chr(13) & _
		"Error description: " & Error$ & Chr(13)  & Chr(13) & _
		"At line: " & Erl & Chr(13)  & Chr(13) & _
		"Date and time: " & Now , 48 ,"An unforeseen error occurred"
	End Select
		
End Function

'____________________________________________________________________________________________________________________________________________________________________________

Sub TestWriteAndReadTextFile

	Dim sResult As String 
	Dim sFilePath As String 
	
	sFilePath = Environ ("HOME") & "239453295/Documents/testfile.txt"
	
	WriteTextFile("This is some text" & Chr(13) & "This is more text on a new line", sFilePath, True, "UTF-8")
	
	sResult = ReadTextFile(sFilePath)
	
	MsgBox sResult 

End Sub 
'____________________________________________________________________________________________________________________________________________________________________________


' LibreOffice Basic's TextOutputStream supports various character encodings. 
' Some common character sets you can use with oTextOutputStream.setEncoding include:

' ISO-8859-1(Latin-1): "ISO-8859-1"
' ISO-8859-13(Baltic Rim): "ISO-8859-13"
' ISO-8859-15(Latin-9): "ISO-8859-15"
' ISO-8859-2(Latin-2): "ISO-8859-2"
' ISO-8859-3(Latin-3): "ISO-8859-3"
' ISO-8859-4(Latin-4): "ISO-8859-4"
' ISO-8859-5(Cyrillic): "ISO-8859-5"
' ISO-8859-6(Arabic): "ISO-8859-6"
' ISO-8859-7(Greek): "ISO-8859-7"
' ISO-8859-8(Hebrew): "ISO-8859-8"
' ISO-8859-9(Turkish): "ISO-8859-9"
' UTF-8: "UTF-8"
' Windows-1250(Central European): "windows-1250"
' Windows-1251(Cyrillic): "windows-1251"
' Windows-1252(Western European): "windows-1252"
' Windows-1253(Greek): "windows-1253"
' Windows-1254(Turkish): "windows-1254"
' Windows-1255(Hebrew): "windows-1255"
' Windows-1256(Arabic): "windows-1256"
' Windows-1257(Baltic): "windows-1257"
' Windows-1258(Vietnamese): "windows-1258"
