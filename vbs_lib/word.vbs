Set fso = CreateObject("Scripting.FileSystemObject")
Set stdin = fso.GetStandardStream(0)
Set stdout = fso.GetStandardStream(1)
Set stderr = fso.GetStandardStream(2)

'to open, parse, search in and scrape from word docx file
'dumping raw data to screen

Class WordFile
	Public oWord
	Public activeDoc
	
	Function Open_Document(doc_file)
		Open_Document = Array(0, "")
	End Function
	
	Function Save_Document(doc_file)
		Save_Document = Array(0, "")
	End Function 
	
	Function Close_Document(doc_file)
		Close_Document = Array(0, "")
	End Function 
	
	Function Get_Word_Application()
		'try catch - if there are no open docs, then warn and quit
		On Error Resume Next
		Err.Clear
		Set oWord = getObject(, "Word.Application")
		If Err.Number = 429 Then
			Get_Word_Application = Array(-1, "There is no open document to work with!")
		Else
			Get_Word_Application = Array(0, "Word application is running.")
		End If
	End Function
	
	Function Assign_Doc(selectedDocIndex)
		Dim index, i
		index = -1
		i = 1
		
		'iterating through documents
		For Each doc In oWord.documents
			If i = CInt(selectedDocIndex) Then	'doc found
				Set activeDoc = doc
				index = i
				Exit For
			Else								'doc not found yet
				i = i + 1
			End If
		Next
		
		'check assigment
		If index <> -1 Then
			activeDoc.Activate
			Assign_Doc = Array(0, "Document '" & activeDoc & "' has been choosen.", activeDoc)
		Else
			Assign_Doc = Array(-1, "Wrong index! No document has been choosen!")
		End If
		'Error: object required -> could be one drive or o365 issue, docx file is not reachable
		'SOLUTION: win restart
		
	End Function
	
	Function Get_Paragraphs_Count()
		Get_Paragraphs_Count = activeDoc.Range.Paragraphs.Count
	End Function
	
	Function Get_Paragraph_Length(index)
		Get_Paragraph_Length = Len(activeDoc.Range.Paragraphs(index).Range.Text)
	End Function
	
	Function Get_Number_Of_Pages()
		Get_Number_Of_Pages = activeDoc.ActiveWindow.Panes(1).Pages.Count
	End Function
	
	Function Read_Paragraph(index)
		If index > 0 And index <= activeDoc.Range.Paragraphs.Count Then
			Dim result
			result = activeDoc.Range.Paragraphs(index).Range.Text
			Read_Paragraph = Array(0, Left(result, Len(result)-1))
		Else
			Read_Paragraph = Array(-1, "Index out of range!")
		End If
	End Function
	
	Function Read_Paragraph_Raw(index)
		If index > 0 And index <= activeDoc.Range.Paragraphs.Count Then
			Read_Paragraph = Array(0, activeDoc.Range.Paragraphs(index).Range.Text)
		Else
			Read_Paragraph = Array(-1, "Index out of range!")
		End If
	End Function
	
	'Special char Chr(13) is added to the input string to write paragraph
	Function Write_Paragraph(index, value)
		If index > 0 And index <= activeDoc.Range.Paragraphs.Count Then
			activeDoc.Range.Paragraphs(index).Range.Text = value & Chr(13)
			Write_Paragraph = Array(0)
		Else
			Write_Paragraph = Array(-1, "Index out of range!")
		End If
	End Function
	
	'No special characters added to the end
	Function Write_Paragraph_Raw(index, value)
		If index > 0 And index <= activeDoc.Range.Paragraphs.Count Then
			activeDoc.Range.Paragraphs(index).Range.Text = value
			Write_Paragraph = Array(0)
		Else
			Write_Paragraph = Array(-1, "Index out of range!")
		End If
	End Function
	
	Function Add_Paragraph(index)
		'
	End Function
	
	Function Delete_Paragraph(index)
		'
	End Function
	
	Function Search_Paragraph_By_Value(value, caseSensitive)
		If value <> "" And (Not isNull(value)) Then
			For i = 1 To activeDoc.Range.Paragraphs.Count
				If StrComp(activeDoc.Range.Paragraphs(i).Range.Text, value, caseSensitive) = 0 Then
					Search_Paragraph_By_Value = Array(0, i)
					Exit Function
				End If
			Next
			Search_Paragraph_By_Value = Array(0, Null)
			Exit Function
		End If
		Search_Paragraph_By_Value = Array(-1)
	End Function
	
	Function Search_Paragraph_By_Style(style)
		If style <> "" And (Not isNull(style)) Then
			For i = 1 To activeDoc.Range.Paragraphs.Count
				If activeDoc.Range.Paragraphs(i).Style = style Then
					Search_Paragraph_By_Style = Array(0, i)
					Exit Function
				End If
			Next
			Search_Paragraph_By_Style = Array(0, Null)
			Exit Function
		End If
		Search_Paragraph_By_Style = Array(-1)
	End Function
	
	Function Search_All_Paragraphs_By_Value(value, caseSensitive)
		If value <> "" And (Not isNull(value)) Then
			Dim result
			Set result = CreateObject("System.Collections.ArrayList")
			'search
			For i = 1 To activeDoc.Range.Paragraphs.Count
				If StrComp(activeDoc.Range.Paragraphs(i).Range.Text, value, caseSensitive) = 0 Then
					result.Add(i)
				End If
			Next
			'examining results
			If result.Count > 0 Then
				Search_All_Paragraphs_By_Value = Array(0, result)
				Exit Function
			Else
				Search_All_Paragraphs_By_Value = Array(0, Null)
				Exit Function
			End If
		End If
		Search_All_Paragraphs_By_Value = Array(-1)
	End Function
	
	Function Search_All_Paragraphs_By_Style(style)
		If style <> "" And (Not isNull(style)) Then
			Dim result
			Set result = CreateObject("System.Collections.ArrayList")
			'search
			For i = 1 To activeDoc.Range.Paragraphs.Count
				If activeDoc.Range.Paragraphs(i).Style = style Then
					result.Add(i)
				End If
			Next
			'examining results
			If result.Count > 0 Then
				Search_All_Paragraphs_By_Style = Array(0, result)
				Exit Function
			Else
				Search_All_Paragraphs_By_Style = Array(0, Null)
				Exit Function
			End If
		End If
		Search_All_Paragraphs_By_Style = Array(-1)
	End Function
	
	Function Search_Paragraph_By_Substr(value, caseSensitive)
		If value <> "" And (Not isNull(value)) Then
			For i = 1 To activeDoc.Range.Paragraphs.Count
				If InStr(1, activeDoc.Range.Paragraphs(i).Range.Text, value, caseSensitive) <> 0 Then
					Search_Paragraph_By_Substr = Array(0, i)
					Exit Function
				End If
			Next
			Search_Paragraph_By_Substr = Array(0, Null)
			Exit Function
		End If
		Search_Paragraph_By_Substr = Array(-1)
	End Function
	
	Function Search_All_Paragraphs_By_Substr(value, caseSensitive)
		If value <> "" And (Not isNull(value)) Then
			Dim result
			Set result = CreateObject("System.Collections.ArrayList")
			'search
			For i = 1 To activeDoc.Range.Paragraphs.Count
				If InStr(1, activeDoc.Range.Paragraphs(i).Range.Text, value, caseSensitive) <> 0 Then
					result.Add(i)
				End If
			Next
			'examining results
			If result.Count > 0 Then
				Search_All_Paragraphs_By_Substr = Array(0, result)
				Exit Function
			Else
				Search_All_Paragraphs_By_Substr = Array(0, Null)
				Exit Function
			End If
		End If
		Search_All_Paragraphs_By_Substr = Array(-1)
	End Function
	
	Sub Dump_Paragraph_By_Index(index)
		stdout.WriteLine("[" & index & "] " & activeDoc.Range.Paragraphs(index).Range.Text)
	End Sub
	
	Sub Dump_Paragraph_By_Range(start_index, end_index)
		For i = start_index To end_index
			If i <= activeDoc.Range.Paragraphs.Count Then
				stdout.WriteLine("[" & i & "] " & activeDoc.Range.Paragraphs(i).Range.Text)
			Else
				stdout.WriteLine("[!] No more paragraphs to print. Too big range given!")
				Exit For
			End If
		Next
	End Sub
	
	Sub Dump_Paragraph_Char_Codes(index)
		Dim paragraphChar, printedChar
		For i = 1 To Len(activeDoc.Range.Paragraphs(index).Range.Text)
			paragraphChar = Mid(activeDoc.Range.Paragraphs(index).Range.Text,i ,1)
			'ASCII printing
			If AscW(paragraphChar) >= 20 And AscW(paragraphChar) <= 126 Then
				printedChar = paragraphChar
				stdout.WriteLine("[" & i & "] " & "'" & printedChar & "' " & AscW(paragraphChar))
			Else
				printedChar = ""
				stdout.WriteLine("[" & i & "] " & "'" & printedChar & "'  " & AscW(paragraphChar)) 'alignment with adding one more space
			End If
		Next
	End Sub
	
	Sub Dump_Paragraph_Char_Codes_By_Range(start_index, end_index)
		For i = start_index To end_index
			stdout.WriteLine("---Paragraph(" & i ")---")
			Dump_Paragraph_Char_Codes(i)
			stdout.WriteLine()
		Next
	End Sub
	
	' *** Table handling functions ***
	
	Function Select_Table(table_index)
		Select_Table = activeDoc.Range.Tables(table_index)
	End Function

	
	'Range means paragraph range
	Function Select_Table_From_Range(range_start, range_end, table_index)
		Select_Table = activeDoc.Range(range_start, range_end).Tables(table_index)
	End Function

	'it needs a paragraph to overwrite it with the table	(for rational reasons the paragraph will be empty 
	'optional parameters for table size and cell size	=>	DefaultTableBehavior, AutoFitBehavior
	Function Create_Empty_Table(table_position_range, num_rows, num_cols)
		activeDoc.Tables.Add(table_position_range, num_rows, num_cols)
		Create_Empty_Table = Array(0, "")
	End Function

	'If we don't want to create a table from zero but want to create a new table that is totally similar to an already existing table
	'We get a table object with Select_Table() and put it into Insert_Existing_Table_To_Paragraph() function
	Function Insert_Existing_Table_To_Paragraph(oTable)
		
	End Function

	' * functions that use a Table object argument *
	
	Function Read_Table(oTable)
		Read_Table = Get_Table_Cells(oTable, 1, oTable.Rows.Count, 1, oTable.Columns.Count)
	End Function

	
	Function Read_Cell_Values_From_Table(oTable, row_start, row_end, col_start, col_end)
		Dim result, tmp
		Set result = CreateObject("System.Collections.ArrayList")
		Set tmp = CreateObject("System.Collections.ArrayList")
		'iterating through table cells
		Dim cell_value
		cell_value = ""
		For i = row_start To row_end
			Set tmp = CreateObject("System.Collections.ArrayList")
			For j = col_start To col_end
				cell_value = oTable.Cell(i, j).Range.Text
				
				'Trim whitespaces 1st round
				If Right(cell_value, 1) = Chr(7) Or Right(cell_value, 1) = Chr(13) Then
					cell_value = Mid(cell_value, 1, Len(cell_value)-1)
				End If
				'Trim whitespaces 2nd round
				If Right(cell_value, 1) = Chr(7) Or Right(cell_value, 1) = Chr(13) Then
					cell_value = Mid(cell_value, 1, Len(cell_value)-1)
				End If
				tmp.Add(cell_value)
			Next
			result.Add(tmp)
		Next
		Read_Cell_Values_From_Table = Array(0, result)
	End Function

	
	Function Write_Table(oTable, values)
		Write_Table = Set_Table_Cells(oTable, 1, oTable.Rows.Count, 1, oTable.Columns.Count, values)
	End Function

	
	Function Write_Cell_Values_To_Table(oTable, row_start, row_end, col_start, col_end, values)
		'iterating through table cells
		For i = row_start To row_end
			For j = col_start To col_end
				cell_value = oTable.Cell(i, j).Range.Text = values(i, j)
			Next
		Next
		Write_Cell_Values_To_Table = Array(0, "")
	End Function

	
	Sub Dump_Table(oTable)
		Call Dump_Cells_From_Table(oTable, 1, oTable.Rows.Count, 1, oTable.Columns.Count)
	End Sub

	
	Sub Dump_Cells_From_Table(oTable, row_start, row_end, col_start, col_end)
		'iterating through table cells
		Dim cell_value
		cell_value = ""
		For i = row_start To row_end
			stdout.Write(" | ")
			For j = col_start To col_end
				cell_value = oTable.Cell(i, j).Range.Text
				
				'Trim whitespaces 1st round
				If Right(cell_value, 1) = Chr(7) Or Right(cell_value, 1) = Chr(13) Then
					cell_value = Mid(cell_value, 1, Len(cell_value)-1)
				End If
				'Trim whitespaces 2nd round
				If Right(cell_value, 1) = Chr(7) Or Right(cell_value, 1) = Chr(13) Then
					cell_value = Mid(cell_value, 1, Len(cell_value)-1)
				End If
				stdout.Write(cell_value)
				stdout.Write(" | ")
			Next
			stdout.WriteLine
		Next
	End Sub


	'instert the new row object to the appropriate position in the rows
	'and returns the now row as well!
	Function Add_Table_Row(oTable, index_to_insert)
		'Rows.Add() has an optional 'beforeRow' parameter with type 'Row' and it is the row index that the new row will be inserted
		Add_Table_Row = Array(0, oTable.Rows.Add(oTable.Rows(index_to_insert)))
	End Function

	Function Add_Table_Row_To_End(oTable)
		'Rows.Add() has an optional 'beforeRow' parameter with type 'Row' and it is the row index that the new row will be inserted
		Add_Table_Row = Array(0, oTable.Rows.Add(oTable.Rows()))
	End Function
	
	Function Add_Table_Column(oTable, index_to_insert)
		'Columns.Add() has an optional 'beforeColumn' parameter with type 'Columns' and it is the column index that the new column will be inserted
		Add_Table_Column = Array(0, oTable.Columns.Add(oTable.Columns(index_to_insert)))
	End Function

	Function Add_Table_Column_To_End(oTable)
		'Columns.Add() has an optional 'beforeColumn' parameter with type 'Columns' and it is the column index that the new column will be inserted
		Add_Table_Column = Array(0, oTable.Columns.Add(oTable.Columns()))
	End Function


	Function Delete_Table_Row(oTable, row_index)
		oTable.Rows(row_index).Delete
		Delete_Table_Row = Array(0, "")
	End Function
	
	Function Delete_All_Rows_In_Table(oTable)
		oTable.Rows().Delete
		Delete_All_Rows_In_Table = Array(0, "")
	End Function


	Function Delete_Table_Column(oTable, col_index)
		oTable.Columns(col_index).Delete
		Delete_Table_Column = Array(0, "")
	End Function

	
	Function Delete_All_Columns_In_Table(oTable)
		oTable.Columns().Delete
		Delete_All_Columns_In_Table = Array(0, "")
	End Function


	Function Delete_Table(oTable)
		oTable.Delete
		Delete_Table = Array(0, "")
	End Function
End Class




