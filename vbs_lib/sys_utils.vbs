Class SysUtils
	Private fso
	
	Sub Class_Initialize()
		Set fso = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	'path is part of the files
	Sub Copy_File(srcfile, dstFile)
		fso.CopyFile(srcFile, dstFile, True)	'3rd parameter => is overwrite enabled
	End Sub
	
	'path is separated
	Sub Copy_File2(src_path, src_file, dst_path, dst_file)
		fso.CopyFile(src_path & src_file, dst_path & dst_file, True)	'3rd parameter => is overwrite enabled
	End Sub
	
	'path is part of the files
	Sub Copy_Files(src_files, dst_files)
		If src_files.Count = dst_files.Count Then
			For i = 0 To src_files.Count-1
				fso.CopyFile(src_files(i), dst_files(i), True)	'3rd parameter => is overwrite enabled
			Next
		Else
			'error handling
		End If
	End Sub
	
	'path is separated
	Sub Copy_Files2(src_path, src_files, dst_path, dst_files)
		If src_files.Count = dst_files.Count Then
			For i = 0 To src_files.Count-1
				fso.CopyFile(src_path & src_files(i), dst_path & dst_files(i), True)	'3rd parameter => is overwrite enabled
			Next
		Else
			'error handling
		End If
	End Sub
	
	Function Run_Shell_Command(cmd)
		Dim oShell, shellResult, result
		Set oShell = CreateObject ("WScript.Shell") 
		'run command
		Set shellResult = oShell.Exec(cmd)

		'wait for command execution
		Do While True
			If shellResult.Status <> 0 Then		'running
				Exit Do
			End If
			Wscript.Sleep(50)
		Loop

		'get output - standard output or error
		Select Case shellResult.Status
			Case 1	'Finished
				result = shellResult.StdOut.ReadAll
			Case 2	'Error
				result = shellResult.StdErr.ReadAll
		End Select
		
		'return output
		result = Mid(result, 1, Len(result)-2)	'Trim CR LF from the end of output
		Run_Shell_Command = result
	End Function
	
	Function Get_Hash_Of_File(filename, algorithm)
		If algorithm = "SHA1" Or algorithm = "SHA256" Or algorithm = "SHA384" Or algorithm = "SHA512" Or algorithm = "MACTripleDES" Or algorithm = "MD5" Or algorithm = "RIPEMD160" Then
			Dim cmd
			cmd = "cmd /c powershell -command " & Chr(34) & "(Get-FileHash '" & filename & "' -Algorithm " & algorithm & " | select hash | Format-Table -HideTableHeaders | Out-String).Trim()" & Chr(34)
			Get_Hash_Of_File = Run_Shell_Command(cmd)
			Exit Function
		Else
			Get_Hash_Of_File = -1
		End If
	End Function
End Class