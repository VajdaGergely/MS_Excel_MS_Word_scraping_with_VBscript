Set fso = CreateObject("Scripting.FileSystemObject")
Set stdin = fso.GetStandardStream(0)
Set stdout = fso.GetStandardStream(1)
Set stderr = fso.GetStandardStream(2)


Class ExcelFile
	Public oExcel
	Public activeWorkbook
	Public activeSheet
	
	' *** Handling Excel Application, Files (Workbooks), Sheets ***
	
	Function Open_Excel_File(filename)
		On Error Resume Next
		  Set oExcel = CreateObject("Excel.Application")
		  oExcel.Visible = True
		  Set activeWorkbook = oExcel.Workbooks.Open(filename)
		  oExcel.Application.Visible = True
		  Open_Excel_File = Array(0, "Excel application is running. Excel file opened.")
		  Exit Function
		If Err Then
			Open_Excel_File = Array(-1, "Error. Can't open Excel application or Excel file!")
			Exit Function
		End If
		On Error Goto 0
	End Function
	
	Function Save_Excel_File()
		activeWorkbook.Save()
		Save_Excel_File = Array(0, "")
	End Function
	
	Function Close_Excel_File()
		activeWorkbook.Close
		'oExcel.Workbooks.Close 	'other excel file may be needed to stay open
		'oExcel.Quit 				'we just want to close the excel file, not the whole excel application

		Set activeSheet = Null
		Set activeWorkbook = Null
		Set oExcel = Null
		Close_Excel_File = Array(0, "")
	End Function
	
	Function Get_Excel_Application()
		'try catch - if there are no open workbooks, then warn and quit
		On Error Resume Next
		Err.Clear
		Set oExcel = getObject(, "Excel.Application")
		If Err.Number = 429 Then
			Get_Excel_Application = Array(-1, "There is no open workbook to work with!")
		Else
			Get_Excel_Application = Array(0, "Excel application is running.")
		End If
	End Function
	
	Function Assign_Workbook(selectedWorkbookIndex)
		Dim index, i
		index = -1
		i = 1
		
		'iterating through workbooks
		For Each workbook In oExcel.Workbooks
			If i = CInt(selectedWorkbookIndex) Then	'workbook found
				Set activeWorkbook = workbook
				index = i
				Exit For
			Else								'workbook not found yet
				i = i + 1
			End If
		Next
		
		'check assigment
		If index <> -1 Then
			activeWorkbook.Activate
			Assign_Workbook = Array(0, "Workbook '" & activeWorkbook.Name & "' has been choosen.", activeWorkbook.Name)
		Else
			Assign_Workbook = Array(-1, "Wrong index! No workbook has been choosen!")
		End If
		'Error: object required -> could be one drive or o365 issue, xlxs file is not reachable
		'SOLUTION: win restart
	End Function

	Function Select_Sheet(selectedSheetIndex)
		Dim index, i
		index = -1
		i = 1
		
		'iterating through workbooks
		For Each sheet In activeWorkbook.Sheets
			If i = CInt(selectedSheetIndex) Then	'sheet found
				Set activeSheet = sheet
				index = i
				Exit For
			Else								'sheet not found yet
				i = i + 1
			End If
		Next
		
		'check assigment
		If index <> -1 Then
			activeSheet.Activate
			Select_Sheet = Array(0, "Sheet '" & activeSheet.Name & "' has been choosen.", activeSheet.Name)
		Else
			Select_Sheet = Array(-1, "Wrong index! No sheet has been choosen!")
		End If
	End Function
	
	
	' *** Cell range I/O functions ***

	Function Clear_Range_Content(range)
		activeSheet.Range(range).ClearContents
		Clear_Range_Content = Array(0, "Cell value has been cleared!")
	End Function
	
	Function Clear_Range_Formatting(range)
		activeSheet.Range(range).ClearFormats
		Clear_Range_Formatting = Array(0, "Cell formatting has been cleared!")
	End Function
	
	Function Clear_Range_Totally(range)
		activeSheet.Range(range).Clear
		Clear_Range_Totally = Array(0, "Cell value and formatting has been cleared!")
	End Function

	Function Clear_Sheet()
		activeSheet.Cells.Clear
		Clear_Sheet = Array(0, "Sheet has been cleared!")
	End Function


	Function Read_Range(range)
		Read_Range = Array(0, activeSheet.Range(range).Value)
		'error handling needed
	End Function

	'array of values for cells one by one
	'works properly with ranges single cells, cols, rows, matrixes e.g. "G10", "G10:G20", "G10:H10", "G10:H20"
	Function Write_Range(range, values)
		Dim i, j	
		'range and values has to be the same element count values
		For i = LBound(values) To UBound(values)
			For j = LBound(values(i)) To UBound(values(i))
				activeSheet.Range(range).Cells(i+1, j+1).Value = values(i)(j)
			Next
		Next
		Write_Range = Array(0, "Cell value(s) have been modified!")
		'error check needed
	End Function
	
	'constant value to all cells
	'works properly with ranges single cells, cols, rows, matrixes e.g. "G10", "G10:G20", "G10:H10", "G10:H20"
	Function Write_Range_With_Constant(range, value)
		activeSheet.Range(range).Value = value
		Write_Range_With_Constant = Array(0, "Cell value(s) has been modified!")
	End Function
	
	'works properly with ranges single cells, cols, rows, matrixes e.g. "G10", "G10:G20", "G10:H10", "G10:H20"
	'the formula is works like the copying of cell formulas in excel
	'the not $ protected fields are gliding down to the next row or column perfectly
	Function Write_Formula_To_Range(range, formula)
		activeSheet.Range(range).Formula = formula
		Write_Formula_To_Range = Array(0, "Cell value(s) has been modified with formula: " & formula)
	End Function


	' *** Cell styling functions ***

	Function Set_Cell_Text_Color(range, color)
		activeSheet.Range(range).Font.Color = RGB(color(0),color(1),color(2))
		Set_Cell_Text_Color = Array(0, "Cell text color has been modified!")
	End Function
	
	Function Set_Cell_Bg_Color(range, color)
		activeSheet.Range(range).Interior.Color = RGB(color(0),color(1),color(2))
		Set_Cell_Bg_Color = Array(0, "Cell background color has been modified!")
	End Function

	Function Get_Cell_Text_Color(range)
		'RGB 0-255 values are handled as a 3 byte binary number and presented at a 3 byte length decimal number
		'its presented in reverse order (B)(G)(R)	=> 255^2 (B)	+	255^1 (G)	+	255^0 (R)
		Dim result, color_dec_value
		result = Array(0, 0, 0)
		color_dec_value = activeSheet.Range(range).Font.Color
		For i = 2 To 0 Step -1
			result(i) = Int(color_dec_value / (256^i))
			color_dec_value = color_dec_value Mod (256^i)
		Next
		Get_Cell_Text_Color = result
	End Function
	
	Function Get_Cell_Bg_Color(range)
		'RGB 0-255 values are handled as a 3 byte binary number and presented at a 3 byte length decimal number
		'its presented in reverse order (B)(G)(R)	=> 255^2 (B)	+	255^1 (G)	+	255^0 (R)
		Dim result, color_dec_value
		result = Array(0, 0, 0)
		color_dec_value = activeSheet.Range(range).Interior.Color
		For i = 2 To 0 Step -1
			result(i) = Int(color_dec_value / (256^i))
			color_dec_value = color_dec_value Mod (256^i)
		Next
		Get_Cell_Bg_Color = result
	End Function

	' *** converting functions between range values and indexes ***
	'e.g.  A1 => 1,1	1,1 => A1	A1:C3 => 1,1,3,3	1,1,3,3 => A1:C3
	

	'	"AB" => int(28)
	Function Cell_Letter_To_Cell_Index(cell_letter)
		Dim result
		result = 0
		cell_letter = StrReverse(cell_letter)
		For i = Len(cell_letter) To 1 Step -1
			result = result + (AscW(UCase(Mid(cell_letter, i, 1)))-64) * (26^(i-1))
		Next
		Cell_Letter_To_Cell_Index = result
		'error handling
	End Function
	
	'	int(28) => "AB"
	Function Cell_Index_To_Cell_Letter(cell_index)
		Dim result, tmp, tmp2
		result = ""
		tmp = 0		'calculated letter value
		tmp2 = 0	'remained amount (from mod) to the next letters
		
		'calc count of letters
		Dim i
		i = 1
		Do While True
			If cell_index < 26 ^ i Then
				Exit Do
			End If
			i = i + 1
		Loop
		
		'calc individual value of letters
		Do While True
			If i > 0 Then
				tmp = Int(cell_index / (26^(i-1)))
				tmp2 = cell_index Mod (26^(i-1))
				result = result & Chr(tmp+64)
				cell_index = tmp2
				i = i - 1
			Else
				Exit Do
			End If
		Loop
		
		'corrigate 0 - 26 - Z problem
		'0 value is not present in excel cols, instead 26 is present as "Z", and the next value is 1 as "A"
		Dim new_char, new_result, overflow
		new_char = ""
		new_result = ""
		overflow = False
		For i = Len(result) To 1 Step -1
			new_char = Asc(Mid(result, i, 1))-64
			
			'handling overflow from previous letter
			If overflow = True Then
				new_char = new_char-1
			End If
			
			'swap invalid excel col chars to the proper valid ones
			If new_char < 1 Then
				If i > 1 Then
					new_char = new_char + 26
					overflow = True
				Else 'if first letter was A and then decreased it won't be "Z" but it is trimed and letter count was lowered by 1
					Exit For
				End If
			Else
				overflow = False
			End If
			
			'char magic
			new_result = new_result & Chr(new_char+64)
		Next
		new_result = StrReverse(new_result)
		Cell_Index_To_Cell_Letter = new_result
		'error handling
	End Function
	
	'	"AB125" => int(28), int(125)
	Function Cell_Value_To_Index_Values(cell_value)
		Dim col, row, tmp, index
		col = ""
		index = 0
		'calc col
		For i = 1 To Len(cell_value)
			tmp = Mid(cell_value, i, 1)
			If ((AscW(tmp) >= 65) And (AscW(tmp) <= 90)) Or ((AscW(tmp) >= 97) And (AscW(tmp) <= 122)) Then
				col = col & tmp
			Else
				index = i
				Exit For
			End If
		Next
		col = Cell_Letter_To_Cell_Index(col) 'converting letters to index value
		
		'error handling
		If col = "" Then
			Cell_Value_To_Index_Values = Array(-1, -1)
			Exit Function
		End If
		
		'calc row
		row = ""
		For i = index To Len(cell_value)
			tmp = Mid(cell_value, i, 1)
			If isNumeric(tmp) Then
				row = row & tmp
			Else
				row = ""
				Exit For
			End If
		Next
		
		'error handling
		If row = "" Then
			Cell_Value_To_Index_Values = Array(-1, -1)
			Exit Function
		End If
		
		Cell_Value_To_Index_Values = Array(col, row)
		'error handling needed
	End Function


	'	int(28), int(125) => "AB125"
	Function Index_Values_To_Cell_Value(col_index, row_index)
		Index_Values_To_Cell_Value = Cell_Index_To_Cell_Letter(col_index) & row_index
	End Function
	
	
	'	"AB125:AE1000" => int(28), int(125), int(31), int(1000)
	Function Range_Value_To_Index_Values(range)
		Dim range_split, result, tmp
		Set result = CreateObject("System.Collections.ArrayList")
		
		'split range to start and end		"A1:D100" => Array("A1", "D100")
		range_split = Split(range, ":")
		
		'get index values (in two array) and union them to a new result array
		tmp = Cell_Value_To_Index_Values(range_split(0))
		result.Add(tmp(0))
		result.Add(tmp(1))
		tmp = Cell_Value_To_Index_Values(range_split(1))
		result.Add(tmp(0))
		result.Add(tmp(1))
		
		Set Range_Value_To_Index_Values = result
		'error handling
	End Function
	
	'	int(28), int(125), int(31), int(1000) => "AB125:AE1000"
	Function Cell_Index_Values_To_Range_Value(col_start, row_start, col_end, row_end)
		Cell_Index_Values_To_Range_Value = Index_Values_To_Cell_Value(col_start, row_start) & ":" & Index_Values_To_Cell_Value(col_end, row_end)
	End Function



	' *** searching functions in ranges ***
	
	'returns array of individual ranges or Null (embedded in Array())
	Function Search_First_Match_In_Range(range, input, isExactMatch, isCaseSensitive)
		Dim result
		
		If isExactMatch = True Then
			isExactMatch = 1
		ElseIf isExactMatch = False Then
			isExactMatch = 2
		End If
		
		'Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
		Set result = activeSheet.Range(range).Find(input, , -4163, isExactMatch, , , isCaseSensitive)
		If Not result Is Nothing Then
			Search_First_Match_In_Range = Array(0, result)
			Exit Function
		Else
			Search_First_Match_In_Range = Array(-1, Null)
			Exit Function
		End If
		'error handling
	End Function

	'returns array of individual ranges or Null (embedded in Array())
	Function Search_All_Matches_In_Range(range, input, isExactMatch, isCaseSensitive)
		Dim result, tmp, firstMatch
		Set result = CreateObject("System.Collections.ArrayList")
		
		If isExactMatch = True Then
			isExactMatch = 1
		ElseIf isExactMatch = False Then
			isExactMatch = 2
		End If
		
		'Find first occurence
		Set tmp = activeSheet.Range(range).Find(input, , -4163, isExactMatch, , , isCaseSensitive)
		Set firstMatch = tmp 'technical stuff, at the end of the search, vbs starts over again, it is needed as an exit point
		'Find more occurences
		If Not tmp Is Nothing Then
			result.Add(tmp)
			Do While True
				Set tmp = activeSheet.Range(range).FindNext(tmp)
				If Not tmp Is Nothing And tmp.Address <> firstMatch.Address Then
					result.Add(tmp)
				Else
					Exit Do
				End If
			Loop
		Else
			Search_All_Matches_In_Range = Array(-1, Null)
			Exit Function
		End If
		Search_All_Matches_In_Range = Array(0, result)
	End Function

	'return array of indexes
	Function Search_Index_Values_Of_First_Match_In_Range(range, input, isExactMatch, isCaseSensitive)
		Dim tmp
		tmp = Search_First_Match_In_Range(range, input, isExactMatch, isCaseSensitive)
		If tmp(0) <> -1 Then
			Search_Index_Values_Of_First_Match_In_Range = Array(0, Array(tmp(1).Column, tmp(1).Row))
			Exit Function
		Else
			Search_Index_Values_Of_First_Match_In_Range = tmp
			Exit Function
		End If
	End Function
	
	'return array of arrays of indexes
	Function Search_Index_Values_Of_All_Matches_In_Range(range, input, isExactMatch, isCaseSensitive)
		Dim tmp
		tmp = Search_All_Matches_In_Range(range, input, isExactMatch, isCaseSensitive)
		If tmp(0) <> -1 Then
			Dim result
			Set result = CreateObject("System.Collections.ArrayList")
			For Each item In tmp(1)
				result.Add(Array(item.Column, item.Row))
			Next
			Search_Index_Values_Of_All_Matches_In_Range = Array(0, result)
			Exit Function
		Else
			Search_Index_Values_Of_All_Matches_In_Range = tmp
			Exit Function
		End If
	End Function
	
	Function Convert_Range_Result_To_ArrayList(oRange)
		Dim result, tmp
		Set result = CreateObject("System.Collections.ArrayList")
		Set tmp = CreateObject("System.Collections.ArrayList")
		For i = 1 To oRange.Rows.Count-1
			Set tmp = CreateObject("System.Collections.ArrayList")
			For j = 1 To oRange.Columns.Count-1
				tmp.Add(oRange.Cells(i, j).Value)
			Next
			result.Add(tmp)
		Next
		Convert_Range_Result_To_ArrayList = Array(0, result)
		'error handling needed
	End Function
	
	Sub Dump_Range_Char_Code(range)
		Dim CellChar, printedChar
		For i = 1 To Len(activeSheet.Range(range).Value)
			CellChar = Mid(activeSheet.Range(range).Value,i ,1)
			'ASCII printing
			If AscW(CellChar) >= 20 And AscW(CellChar) <= 126 Then
				printedChar = CellChar
				stdout.WriteLine("[" & i & "] " & "'" & printedChar & "' " & AscW(CellChar))
			Else
				printedChar = ""
				stdout.WriteLine("[" & i & "] " & "'" & printedChar & "'  " & AscW(CellChar)) 'alignment with adding one more space
			End If
		Next
	End Sub
End Class


