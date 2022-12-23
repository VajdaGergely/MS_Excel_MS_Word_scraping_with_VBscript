Class FileIo
	Function Read_File(filename, ByRef output)
		Dim file, content
		Set file = fso.OpenTextFile(filename, 1)
		Dim line
		Do Until file.AtEndOfStream
			line = file.ReadLine
			output.Add(line)
		Loop
		file.Close
		Read_File = 0
	End Function
	
	'
	'using array as input
	Function Write_File(filename, content)
		'error kezeles kene bele majd
		Dim file
		Set file = fso.OpenTextFile(filename, 2)
		
		For Each row In content
			file.WriteLine row
		Next
		file.Close
		Write_File = 0
	End Function
End Class