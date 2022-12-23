Set fso = CreateObject("Scripting.FileSystemObject")
Set stdin = fso.GetStandardStream(0)
Set stdout = fso.GetStandardStream(1)
Set stderr = fso.GetStandardStream(2)

ExecuteGlobal fso.OpenTextFile("..\vbs_lib\file_utils.vbs").ReadAll()



Class CliIo
	Sub Print_Empty()
		stdout.WriteLine()
	End Sub

	Sub Print(input)
		stdout.WriteLine(input)
	End Sub

	Sub Print_NoLf(input)
		stdout.Write(input)
	End Sub

	Function Read_Input()
		Read_Input = stdin.ReadLine
	End Function

	Function Read_Input2(message)
		stdout.Write(message)
		Read_Input2 = stdin.ReadLine
	End Function

	Sub Read_InputList(object_name, ByRef result)
		Dim count
		Set result = CreateObject("System.Collections.ArrayList")
		count = Read_Input2(object_name & " count: ")
		For i = 0 To count-1
			result.Add(Read_Input2(object_name & "(" & i+1 & "): "))
		Next
	End Sub

	Sub Print_Ok(input)
		stdout.WriteLine("[+] " & input)
	End Sub

	Sub Print_Nok(input)
		stdout.WriteLine("[-] " & input)
	End Sub

	Sub Print_Inf(input)
		stdout.WriteLine("[*] " & input)
	End Sub

	Sub Print_Warn(input)
		stdout.WriteLine("[!] " & input)
	End Sub

	Sub Print_Result(result)
		Print_Empty
		If result(0) <> -1 Then
			Print_Ok(result(1))
		Else
			Print_Nok(result(1))
		End If
		Print_Empty
	End Sub

	Sub Print_Message(msgText, msgType)
		Print_Empty
		If msgType = "ok" Then
			Print_Ok(msgText)
		ElseIf msgType = "nok" Then
			Print_Nok(msgText)
		ElseIf msgType = "inf" Then
			Print_Inf(msgText)
		ElseIf msgType = "warn" Then
			Print_Warn(msgText)
		End If
		Print_Empty
	End Sub

	Sub Print_Header(input)
		stdout.WriteLine("### " & input & " ###")
	End Sub

	Sub Print_SubHeader(input)
		stdout.WriteLine("---" & input & "---")
	End Sub

	'print a submenu with options and returns the selected index
	'it changes <option> string to the text of the selected option in the success message
	Function Choose_Option(header, options, input_text, error_msg, success_msg)
		Dim cmd
		cmd = ""
		
		Do While True
			Print_SubHeader(header)
			'print options
			For i = 0 To options.Count-1
				Print("(" & i+1 & ") " & options(i))
			Next
			
			'read user input
			cmd = cli.Read_Input2(input_text & ": ")
			Print_Empty
			
			'if input is valid, return command index
			If isNumeric(cmd) And CInt(cmd) > 0 And CInt(cmd) <= options.Count Then
				success_msg = Replace(success_msg, "<option>", options(cmd-1))
				Print_Ok(success_msg)
				Print_Empty
				Exit Do
			Else
				Print_Nok(error_msg)
				Print_Empty
			End If
		Loop
		Choose_Option = cmd
	End Function

	Function Choose_Option2()
		'
	End Function

	Function Set_Parameter(description, parameter)
		stdout.WriteLine description & parameter
		stdout.Write "New value: "
		Dim input
		input = stdin.ReadLine
		If input <> "" Then
			parameter = input
		End If
		Set_Parameter = parameter
	End Function

	Sub Press_Enter()
		stdout.WriteLine()
		stdout.Write("Press ENTER to continue")
		stdin.ReadLine()
		stdout.WriteLine()
	End Sub
End Class


Class CliConfig
	Private globalParameters
	Private config_filename
	Private file_io
	Private cli_io
	
	Sub Class_Initialize()
		Set globalParameters = CreateObject("System.Collections.ArrayList")
		Set file_io = New FileIO
		Set cli_io = New CliIO
	End Sub
	
	'This sub creates the structure of the parameters list, Load_Config only modifies the values by local file
	Sub Add_Config_Parameter(name, description, value)
		globalParameters.Add(Array(name, description, value))
	End Sub
	
	Sub Set_Config_Filename(filename)
		config_filename = filename
	End Sub
	
	Function Get_Para_Val(name)
		For Each para In globalParameters
			If para(0) = name Then
				Get_Para_Val = para(2)
				Exit Function
			End If
		Next
		Get_Para_Val = -1
	End Function
	
	'Add_Config needs to be called before
	'
	Sub Load_Config()
		'if config file is not exist, or not contains the needed lines, then maybe errors occur
		Dim result, input
		Set input = CreateObject("System.Collections.ArrayList")
		result = file_io.Read_File(config_filename, input)
		'error handling, if the size is not equal (dict and file row count), then error occur
		'If input.Count = globalParameters.Count
		For i = 0 To globalParameters.Count - 1
			globalParameters(i) = Array(globalParameters(i)(0), globalParameters(i)(1), input(i))
		Next
		'Else
			'error....
		'End If
	End Sub
	
	
	Private Function Save_Config()
		'error kezeles kene bele majd
		'Need to use custom file_handling code, because values are stored in multi dimension array!!!
		Dim file
		Set file = fso.OpenTextFile(config_filename, 2)
		For i = 0 To globalParameters.Count - 1
			file.WriteLine globalParameters(i)(2)
		Next
		file.Close
		Save_Config = 0
	End Function
	
	Sub Set_Config()
		'load parameters
		Call Load_Config()
		
		'modify or skip parameters
		For i = 0 To globalParameters.Count - 1
			globalParameters(i) = Array(globalParameters(i)(0), globalParameters(i)(1), cli_io.Set_Parameter(globalParameters(i)(1) & ": ", globalParameters(i)(2)))
		Next
		
		'save parameters
		result = Save_Config()
	End Sub
	
	Sub Print_Config()
		'load parameters
		Call Load_Config()
		
		'print parameters
		For Each para In globalParameters
			cli_io.Print(para(0) & ": " & para(2))
		Next
	End Sub
	
	Sub Reset_Config(values)
		'reset parameters
		For i = 0 To globalParameters.Count - 1
			globalParameters(i) = Array(globalParameters(i)(0), globalParameters(i)(1), values(i))
		Next
		
		'save parameters
		result = Save_Config()
		
		'print info
		For Each para In globalParameters
			cli_io.Print(para(1) & " => '" & para(2) & "'")
		Next
		'cli_io.Print_Ok("Parameters has been reset.")
	End Sub
End Class

Class CliMenu
	Private commandList
	
	Sub Class_Initialize()
		Set commandList = CreateObject("System.Collections.ArrayList")
	End Sub
	
	Sub AddCmd(description, func)
		commandList.Add (Array(description, func))
	End Sub
	
	Sub Press_Enter()
		stdout.WriteLine()
		stdout.Write("Press ENTER to continue")
		stdin.ReadLine()
		stdout.WriteLine()
	End Sub
	
	Sub Run()
		Dim input, index
		Do While True
			'print command list
			stdout.WriteLine("---Commands---")
			For i = 0 To commandList.Count - 1
				stdout.WriteLine("(" & i + 1 & ") " & commandList(i)(0))
			Next
			stdout.WriteLine "(0) Exit"
			
			'get command to run
			stdout.Write "Command: "
			input = stdin.ReadLine
			stdout.WriteLine()
			
			'validating english keyboard layout
			If input = "`" Then
				input = "0"
			End If	
			
			'exiting
			If input = "0" Then
				stdout.WriteLine()
				stdout.WriteLine "Exiting..."
				Exit Do
			End If
			
			input = input - 1 'Array indexed from '0' but commands counted from '1'
			
			'search and run command
			If isNumeric(input) And CInt(input) >= 0 And CInt(input) < commandList.Count Then
				'There are differences in the syntax of function pointer usage
				If Instr(1, commandList(input)(1), ".", False) <> 0 Then
					'if external lib contains the function to call (lib_name.funcname) (contains a "." char)
					'then 'Eval' is used
					Eval(commandList(input)(1))
				Else
					'if the calling code contains the function (that instantiated the CliMenu Class)
					'then 'GetRef' is used
					Dim func
					Set func = GetRef(commandList(input)(1))
					Call func
				End If
			Else
				stdout.WriteLine("Wrong command!")
			End If
			
			Press_Enter()
		Loop
	End Sub
End Class