Set fso = CreateObject("Scripting.FileSystemObject")
Set stdin = fso.GetStandardStream(0)
Set stdout = fso.GetStandardStream(1)
Set stderr = fso.GetStandardStream(2)

Execute fso.OpenTextFile("..\vbs_lib\word.vbs").ReadAll()
Execute fso.OpenTextFile("..\vbs_lib\cli_utils.vbs").ReadAll()

'wrappers to word.vbs subs and functions
'to add I/O console things
'to add cli things
'to pretty printing queried results

'plus utilities
'caching paragraphs and paragraph attributes
'easy to handle index-arrays of important elements (like headings)



Class WordUtils
	Public word_file
	Private cli
	
	Private heading_styles
	Public headings_list		'ArrayList of ArrayList of Array		headings grouped by lvl		for getting headings of specific lvl
	
	Sub Class_Initialize()
		Set word_file = New WordFile
		Set cli = New CliIo
		Set headings_list = CreateObject("System.Collections.ArrayList")
		Set heading_styles = CreateObject("System.Collections.ArrayList")
	End Sub
	
	
	Function Attach_Document()
		Dim result, selectedDocIndex
		result = word_file.Get_Word_Application()
		If result(0) = -1 Then
			cli.Print_Result(result)
		Else
			Do While True
				'list open docs
				cli.Print_SubHeader("Currently open documents")
				For i = 1 To word_file.oWord.documents.Count
					cli.Print "(" & i & ") " & word_file.oWord.documents(i)
				Next
				
				'select doc to work with
				cli.Print_Empty
				selectedDocIndex = cli.Read_Input2("Selected document: ")
				
				'assigning doc
				If isNumeric(selectedDocIndex) Then		'validating user input and 
					If selectedDocIndex = "`" Then 		'validating english keyboard layout
						selectedDocIndex = "0"
					End If
					
					result = word_file.Assign_Doc(selectedDocIndex)	'assign doc
					cli.Print_Result(result)
					If result(0) <> -1 Then
						word_file.activeDoc.Activate
						Exit Do
					End If
				Else
					Call Print_Message("Wrong input value!", "nok")
				End If
			Loop
		End If
		Attach_Document = result
	End Function
	
	Sub Set_Heading_Styles(style_input_list)
		Set heading_styles = CreateObject("System.Collections.ArrayList")
		For Each item In style_input_list
			heading_styles.Add(item)
		Next
	End Sub
	
	Sub Set_Heading_Styles_From_Cli()
		Dim tmp, heading_count
		Set heading_styles = CreateObject("System.Collections.ArrayList")
		heading_count = cli.Read_Input2("Count of headings: ")
		For i = 0 To heading_count-1
			tmp = cli.Read_Input2("Value of heading(" & i+1 & "): ")
			heading_styles.Add(tmp)
		Next
	End Sub
	
	'headings_list	=>	ArrayList of ArrayList of Array		headings grouped by level
	Sub Read_All_Headings()
		If heading_styles.Count > 0 Then
			cli.Print_Inf("This could take a few seconds...")
			cli.Print_Empty
			
			Dim lvl_counter_list 'heading level counters
			Set lvl_counter_list = CreateObject("System.Collections.ArrayList")
			Set headings_list = CreateObject("System.Collections.ArrayList")
			
			'init heading list, creating array lists for every heading level
			'init lvl_counter_list
			Dim tmp
			For i = 0 To heading_styles.Count-1
				lvl_counter_list.Add(0)
				Set tmp = CreateObject("System.Collections.ArrayList")
				headings_list.Add(tmp)
			Next
			
			'fill heading lists
			'iterating through paragraphs
			For i = 1 To word_file.activeDoc.Range.Paragraphs.Count
				'iterating through heading styles
				For j = 0 To heading_styles.Count-1
					If word_file.activeDoc.Range.Paragraphs(i).Style = heading_styles(j) Then
						'increasing the appropriate heading list level value
						lvl_counter_list(j) = lvl_counter_list(j) + 1
						
						'set all child heading list level value to null
						For k = j+1 To heading_styles.Count-1
							lvl_counter_list(k) = 0
						Next
						
						'reset tmp
						Set tmp = CreateObject("System.Collections.ArrayList") 'it will contains the structure that contain all details about one heading entry
						
						'add paragraph index "i" to tmp structure
						tmp.Add(i)
						
						'add previous heading list level values to heading list
						For k = 0 To j-1 'the end is j-1 because j is the current item, and we just iterating on parent items
							tmp.Add(lvl_counter_list(k)) 'lvl_counter_list contains the parent heading level values
						Next
						
						'add the whole heading structure to the appropiate level of heading list "j"
						headings_list(j).Add(tmp)
						
						'skip the upcoming search on the actual paragraph and move to the next paragraph
						Exit For
					End If
				Next
			Next
			cli.Print_Ok("Headings has been loaded.")
		Else
			cli.Print_Nok("No heading style has been set yet!")
		End If
	End Sub
	
	
	'Dump headings from previously read and stored list of headings
	Sub Dump_All_Headings_By_Group()
		If heading_styles.Count > 0 Then
			If headings_list.Count > 0 Then
				For i = 0 To headings_list.Count-1
					cli.Print_Subheader("Lvl " & i+1)
					For Each row In headings_list(i)
							cli.Print_NoLf("[" & row(0) & "] ")
							cli.Print_NoLf(word_file.activeDoc.Range.Paragraphs(row(0)).Range.Text)
						cli.Print()
					Next
					cli.Print_Empty
				Next
			Else
				cli.Print_Nok("No headings loaded yet!")
			End If
		Else
			cli.Print_Nok("No heading style has been set yet!")
		End If
	End Sub
	
	Function Get_All_Headings_By_Lvl(level)
		Dim result
		Set result = CreateObject("System.Collections.ArrayList")
		For Each row In headings_list(level-1)
			result.Add(row(0))
		Next
		Set Get_All_Headings_By_Lvl = result
	End Function
	
	Sub Print_Headings_By_Lvl()
		If headings_list.Count > 0 Then
			Dim result
			'get headings
			Set result = Get_All_Headings_By_Lvl(cli.Read_Input2("Level: "))
			
			'print headings
			cli.Print_Subheader("Lvl " & level+1)
			For Each value In result
					cli.Print_NoLf("[" & value & "] ")
					cli.Print_NoLf(word_file.activeDoc.Range.Paragraphs(value).Range.Text)
				cli.Print()
			Next
			cli.Print_Empty
		Else
			cli.Print_Nok("No headings loaded yet!")
		End If
	End Sub
	
	Function Get_Headings_By_Lvl_And_Parent(level, parent_level, parent_index)
		Dim result
		Set result = CreateObject("System.Collections.ArrayList")
		If level >= 1 And parent_level >= 0 And parent_level < level And parent_index >= 0 Then 'level cant be 0, because 0 has no parent
			For Each row In headings_list(level)
				If row(parent_level+1) = parent_index+1 Then	'+1 is needed in the index, because index 0 is the paragraph index (parent_index+1 is strange...)
					result.Add(row(0))
				End If
			Next
		Else
			result = Nothing
		End If
		Set Get_Headings_By_Lvl_And_Parent = result
	End Function
	
	Sub Print_Headings_By_Lvl_And_Parent()
		If headings_list.Count > 0 Then
			Dim result, level, parent_level, parent_index
			level = CInt(cli.Read_Input2("Level: ")-1)
			parent_level = CInt(cli.Read_Input2("Parent level: ")-1)
			parent_index = CInt(cli.Read_Input2("Parent index: ")-1)
			
			If level < 1 Then
				cli.Print_Nok("Level has to be larger than 1!")
				Exit Sub
			End If
			
			'get headings
			Set result = Get_Headings_By_Lvl_And_Parent(level, parent_level, parent_index)
			
			'print headings
			cli.Print_Subheader("Lvl " & level+1)
			For Each value In result
					cli.Print_NoLf("[" & value & "] ")
					cli.Print_NoLf(word_file.activeDoc.Range.Paragraphs(value).Range.Text)
				cli.Print()
			Next
			cli.Print_Empty
		Else
			cli.Print_Nok("No headings loaded yet!")
		End If
	End Sub
	
	'Pretty print headings
	'ArrayList of Array		headings sorted by index, no grouping
	'Reads all headings and prints them with identation
	Sub Dump_All_Headings_Raw()
		'identation
		'pretty print on the fly
		'no storing values
		'need to be read again
		
		If heading_styles.Count > 0 Then
			cli.Print_Inf("This could take a few seconds...")
			cli.Print_Empty
			
			Dim ident
			ident = "  "
			
			'iterating through paragraphs
			For i = 1 To word_file.activeDoc.Range.Paragraphs.Count
				'iterating through heading styles
				For j = 0 To heading_styles.Count-1
					If word_file.activeDoc.Range.Paragraphs(i).Style = heading_styles(j) Then
						'print index
						cli.Print_NoLf("[" & i & "] ")
						'print identation
						For k = 0 To j-1
							cli.Print_NoLf(ident)
						Next			
						'print paragraph
						cli.Print_NoLf(word_file.activeDoc.Range.Paragraphs(i).Range.Text)
						cli.Print()
						
						'skip the upcoming search on the actual paragraph and move to the next paragraph
						Exit For
					End If
				Next
			Next
		Else
			cli.Print_Nok("No heading style has been set yet!")
		End If
	End Sub
	
	Sub Print_Style_Of_Paragraph()
		Dim index
		index = cli.Read_Input2("index: ")
		cli.Print_Empty
		cli.Print("[" & index & "] " & word_file.activeDoc.Range.Paragraphs(index).Range.Text)
		cli.Print("[Style] " & word_file.activeDoc.Range.Paragraphs(index).Style)
	End Sub
	
	Private Function Is_Case_Sensitive()
		Dim input
		Do While True
			input = cli.Read_Input2("case sensitive [Y]es, [N]o (default): ")
			If input = "N" Or input = "n" Or input = "" Then
				Is_Case_Sensitive = 1
				Exit Function
			ElseIf input = "Y" Or input = "y" Then
				Is_Case_Sensitive = 0
				Exit Function
			End If
		Loop
	End Function
	
	Sub Search_Paragraph_By_Value()
		Dim result, value
		value = cli.Read_Input2("value: ")
		result = word_file.Search_Paragraph_By_Value(value & Chr(13), Is_Case_Sensitive())
		cli.Print_SubHeader("Search Paragraph By Value")
		If result(0) <> -1 Then
			If result(0) = 0 And (Not isNull(result(1))) Then
				temp = word_file.activeDoc.Range.Paragraphs(result(1)).Range.Text
				cli.Print("[" & result(1) & "] '" & Left(temp, Len(temp)-1) & "'")
			Else
				cli.Print_Inf("No paragraphs found with value '" & value &"'!")
			End If
		Else
			cli.Print_Nok("An error occured!")
		End If
	End Sub
	
	Sub Search_Paragraph_By_Style()
		Dim result, style
		style = cli.Read_Input2("style: ")
		result = word_file.Search_Paragraph_By_Style(style)
		cli.Print_SubHeader("Search Paragraph By Style")
		If result(0) <> -1 Then
			If result(0) = 0 And (Not isNull(result(1))) Then
				temp = word_file.activeDoc.Range.Paragraphs(result(1)).Range.Text
				cli.Print("[" & result(1) & "] '" & Left(temp, Len(temp)-1) & "'")
			Else
				cli.Print_Inf("No paragraphs found with style '" & style &"'!")
			End If
		Else
			cli.Print_Nok("An error occured!")
		End If
	End Sub
	
	Sub Search_All_Paragraphs_By_Value()
		Dim result, value
		value = cli.Read_Input2("value: ")
		result = word_file.Search_All_Paragraphs_By_Value(value & Chr(13), Is_Case_Sensitive())
		cli.Print_SubHeader("Search Paragraph By Value")
		If result(0) <> -1 Then
			If result(0) = 0 And (Not isNull(result(1))) Then
				For i = 0 To result(1).Count - 1
					temp = word_file.activeDoc.Range.Paragraphs(result(1)(i)).Range.Text
					cli.Print("[" & result(1)(i) & "] '" & Left(temp, Len(temp)-1) & "'")
				Next
			Else
				cli.Print_Inf("No paragraphs found with value '" & value &"'!")
			End If
		Else
			cli.Print_Nok("An error occured!")
		End If
	End Sub
	
	Sub Search_All_Paragraphs_By_Style()
		Dim result, style
		style = cli.Read_Input2("style: ")
		result = word_file.Search_All_Paragraphs_By_Style(style)
		cli.Print_SubHeader("Search Paragraph By Style")
		If result(0) <> -1 Then
			If result(0) = 0 And (Not isNull(result(1))) Then
				Dim temp 'placeholder for the trimmed and printed value
				temp = ""
				For i = 0 To result(1).Count - 1
					'even if the style has been searched, we want to print the value of the paragraphs, (without the trailing Chr(13))
					temp = word_file.activeDoc.Range.Paragraphs(result(1)(i)).Range.Text
					cli.Print("[" & result(1)(i) & "] '" & Left(temp, Len(temp)-1) & "'")
				Next
			Else
				cli.Print_Inf("No paragraphs found with style '" & style &"'!")
			End If
		Else
			cli.Print_Nok("An error occured!")
		End If
	End Sub
	
	Sub Search_Paragraph_By_Substr()
		Dim result, value
		value = cli.Read_Input2("value: ")
		result = word_file.Search_Paragraph_By_Substr(value, Is_Case_Sensitive())
		cli.Print_SubHeader("Search Paragraph By Value")
		If result(0) <> -1 Then
			If result(0) = 0 And (Not isNull(result(1))) Then
				Dim temp 'placeholder for the trimmed and printed value
				temp = word_file.activeDoc.Range.Paragraphs(result(1)).Range.Text
				cli.Print("[" & result(1) & "] '" & Left(temp, Len(temp)-1) & "'")
			Else
				cli.Print_Inf("No paragraphs found with value '" & value &"'!")
			End If
		Else
			cli.Print_Nok("An error occured!")
		End If
	End Sub
	
	Sub Search_All_Paragraphs_By_Substr()
		Dim result, value
		value = cli.Read_Input2("value: ")
		result = word_file.Search_All_Paragraphs_By_Substr(value, Is_Case_Sensitive())
		cli.Print_SubHeader("Search Paragraph By Value")
		If result(0) <> -1 Then
			If result(0) = 0 And (Not isNull(result(1))) Then
				Dim temp 'placeholder for the trimmed and printed value
				temp = ""
				For i = 0 To result(1).Count - 1
					temp = word_file.activeDoc.Range.Paragraphs(result(1)(i)).Range.Text
					cli.Print("[" & result(1)(i) & "] '" & Left(temp, Len(temp)-1) & "'")
				Next
			Else
				cli.Print_Inf("No paragraphs found with value '" & value &"'!")
			End If
		Else
			cli.Print_Nok("An error occured!")
		End If
	End Sub
	
	
	Sub Dump_Paragraph_By_Index()
		Dim index
		index = cli.Read_Input2("index: ")
		cli.Print_SubHeader("Paragraphs By Index")
		Call word_file.Dump_Paragraph_By_Index(index)
		cli.Print_Empty
	End Sub
	
	Sub Dump_Paragraph_By_Range()
		Dim start_index, end_index
		start_index = cli.Read_Input2("start: ")
		end_index = cli.Read_Input2("end: ")
		cli.Print_SubHeader("Paragraphs By Range")
		cli.Print("range: " & start_index & "-" & end_index)
		cli.Print_Empty
		Call word_file.Dump_Paragraph_By_Range(start_index, end_index)
		cli.Print_Empty
	End Sub
	
	Sub Dump_Paragraph_Char_Codes()
		Dim index
		index = cli.Read_Input2("index: ")
		cli.Print_SubHeader("Paragraph length")
		cli.Print(word_file.Get_Paragraph_Length(index))
		cli.Print_SubHeader("Paragraph char codes")
		Call word_file.Dump_Paragraph_Char_Codes(index)
		cli.Print_Empty
	End Sub
	
	Sub Dump_Paragraph_Char_Codes_By_Range()
		Dim start_index, end_index
		start_index = cli.Read_Input2("start index: ")
		end_index = cli.Read_Input2("end index: ")
		Call word_file.Dump_Paragraph_Char_Codes_By_Range(start_index, end_index)
		cli.Print_Empty
	End Sub
End Class