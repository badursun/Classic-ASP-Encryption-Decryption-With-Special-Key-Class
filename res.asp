<%
Class RabbitEncryptionSystem 'RES
'######################################################################
' RES - Rabbit Encryption System
' Special Thanks: 
' 	RC4 Encryption Using ASP & VBScript By Mike Shaffer
' 	Encrypt and decrypt functions for classic ASP (by TFI)
' 	This class contains specially compiled and enhanced functions.
' 	https://github.com/badursun/Classic-ASP-Encryption-Decryption-With-Special-Key-Class/
'######################################################################
	Private g_KeyLen, g_KeyLocation, g_DefaultKey, cryptkey, SavePathLocation, fileSaveState, LastSavedFileName

'---------------------------------------
' Class Initialize
'---------------------------------------
	Private Sub Class_Initialize()
		fileSaveState 		= False
		LastSavedFileName 	= ""
		g_KeyLocation 		= "../keyFiles"
		g_KeyLen 			= 512
		g_DefaultKey 		= "GNQ?4i0-*\CldnU+[vrF1j1PcWeJfVv4QGBurFK6}*l[H1S:oY\v@U?i" &_
						  ",oD]f/n8oFk6NesH--^PJeCLdp+(t8SVe:ewY(wR9p-CzG<,Q/(U*.pX" &_ 
						  "Diz/KvnXP`BXnkgfeycb)1A4XKAa-2G}74Z8CqZ*A0P8E[S`6RfLwW+P" &_ 
						  "c}13U}_y0bfscJ<vkA[JC;0mEEuY4Q,([U*XRR}lYTE7A(O8KiF8>W/m" &_
						  "1D*YoAlkBK@`3A)trZsO5xv@5@MRRFkt\"

		cryptkey 			= g_DefaultKey
		SavePathLocation	= Server.Mappath( g_KeyLocation )
	End Sub
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Class Terminate
'---------------------------------------
	Private Sub Class_Terminate()

	End Sub
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Get Default Key
'---------------------------------------
	Public Property Get KeySavePath()
		KeySavePath = SavePathLocation
	End Property
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Get Default Key
'---------------------------------------
	Public Property Get DefaultKey()
		DefaultKey = g_DefaultKey
	End Property
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Let cryptkey
'---------------------------------------
	Public Property Let FilePath(vVal)
		If Len(vVal) < 1 Then 
			SavePathLocation = Server.Mappath( g_KeyLocation )
		Else 
			SavePathLocation = Server.Mappath( vVal )
		End If
	End Property
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Let cryptkey
'---------------------------------------
	Public Property Let Key(vVal)
		If Len(vVal) < 1 Then 
			cryptkey = KeyGeN( g_KeyLen )
		Else 
			cryptkey = vVal
		End If
	End Property
	Public Property Get Key()
		Key = cryptkey
	End Property
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Let File Saved
'---------------------------------------
	Public Property Get LastSavedFile()
		LastSavedFile = LastSavedFileName
	End Property
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Let File Saved
'---------------------------------------
	Public Property Get FileSaved()
		FileSaved = fileSaveState
	End Property
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Write Key File Pyhsical Folder
'---------------------------------------
	Public Function WriteKeyToFile(strFileName)
		LastSavedFileName = ""
		On Error Resume Next

		Dim keyFile, fso
		set fso = Server.CreateObject("scripting.FileSystemObject") 
		set keyFile = fso.CreateTextFile(SavePathLocation&"/"&strFileName, true) 
			keyFile.WriteLine( cryptkey )
			keyFile.Close
		
		If Err <> 0 Then
			fileSaveState = False
		Else
			fileSaveState = True
			LastSavedFileName = strFileName
		End If

		WriteKeyToFile = fileSaveState

		set keyFile = Nothing
		set fso = Nothing
	End Function
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Read Key From File 
'---------------------------------------
	Public Function ReadKeyFromFile(strFileName)
		Dim keyFile, Fso, f
		Set Fso = Server.CreateObject("Scripting.FileSystemObject") 
		
		tmpFileFullPath = SavePathLocation&"/"&strFileName

		If Not Fso.FileExists( tmpFileFullPath ) = True Then
			ReadKeyFromFile = "File Not Found"
		Else
			Set f = Fso.GetFile( tmpFileFullPath ) 
			Set ts = f.OpenAsTextStream(1, -2)

			Do While not ts.AtEndOfStream
				keyFile = keyFile & ts.ReadLine
			Loop 

			ReadKeyFromFile =  keyFile
		End If

		Set ts = Nothing 
		Set f = Nothing 
		Set Fso = Nothing
	End Function
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' URL Encode/Decode
'---------------------------------------
	Private Function URLDecode4Encrypt(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I
		If IsNull(sConvert) Then
			URLDecode4Encrypt = ""
			Exit Function
		End If
		
		'sOutput = REPLACE(sConvert, "+", " ") ' convert all pluses to spaces
		sOutput=sConvert
		aSplit = Split(sOutput, "%")
		If IsArray(aSplit) Then
			sOutput = aSplit(0)
			For I = 0 to UBound(aSplit) - 1
				sOutput = sOutput &  Chr("&H" & Left(aSplit(i + 1), 2)) & Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
			Next
		End If
		URLDecode4Encrypt = sOutput
	End Function
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' KeyGen generator
'---------------------------------------
	Private Function KeyGeN(iKeyLength)
		Dim k, iCount, strMyKey
		lowerbound = 35 
		upperbound = 96
		Randomize
		for i = 1 to iKeyLength
			s = 255
			k = Int(((upperbound - lowerbound) + 1) * Rnd + lowerbound)
			strMyKey =  strMyKey & Chr(k) & ""
		next
		KeyGeN = strMyKey
	End Function
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Encrypt
'---------------------------------------
	Public function Encrypt(inputstr)
		Dim i,x
		outputstr=""
		cc=0
		for i=1 to len(inputstr)
			x=asc(mid(inputstr,i,1))
			x=x-48
			if x<0 then x=x+255
			x=x+asc(mid(cryptkey,cc+1,1))
			if x>255 then x=x-255
			outputstr=outputstr&chr(x)
			cc=(cc+1) mod len(cryptkey)
		next
		Encrypt = server.urlencode(replace(outputstr,"%","%25"))
	end function
'---------------------------------------
' END
'---------------------------------------

'---------------------------------------
' Decrypt
'---------------------------------------
	Public function Decrypt(byval inputstr)
		Dim i,x
		inputstr=URLDecode4Encrypt(inputstr)
		outputstr=""
		cc=0
		for i=1 to len(inputstr)
			x=asc(mid(inputstr,i,1))
			x=x-asc(mid(cryptkey,cc+1,1))
			if x<0 then x=x+255
			x=x+48
			if x>255 then x=x-255
			outputstr=outputstr&chr(x)
			cc=(cc+1) mod len(cryptkey)
		next
		Decrypt = outputstr
	end function
'---------------------------------------
' END
'---------------------------------------

End Class
%>
