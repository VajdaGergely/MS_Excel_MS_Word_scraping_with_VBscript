Set fso = CreateObject("Scripting.FileSystemObject")
Set stdin = fso.GetStandardStream(0)
Set stdout = fso.GetStandardStream(1)
Set stderr = fso.GetStandardStream(2)

Execute fso.OpenTextFile("..\vbs_lib\excel.vbs").ReadAll()
Execute fso.OpenTextFile("..\vbs_lib\word_utils.vbs").ReadAll()
'Execute fso.OpenTextFile("..\vbs_lib\cli_utils.vbs").ReadAll()


Class ExcelUtils
	Public excel_file
	Private cli

	Sub Class_Initialize()
		Set excel_file = New ExcelFile
		Set cli = New CliIo
	End Sub

	Function Attach_Workbook()
		Dim result, selectedWorkbookIndex
		result = excel_file.Get_Excel_Application()
		If result(0) = -1 Then
			cli.Print_Result(result)
		Else
			Do While True
				'list open docs
				cli.Print_SubHeader("Currently open workbooks")
				Dim i
				i = 1
				For Each workbook In excel_file.oExcel.Workbooks
					cli.Print("(" & i & ") " & workbook.Name)
					i = i + 1
				Next
				
				'select doc to work with
				cli.Print_Empty
				selectedWorkbookIndex = cli.Read_Input2("Selected workbook: ")
				
				'assigning doc
				If isNumeric(selectedWorkbookIndex) Then		'validating user input and 
					If selectedWorkbookIndex = "`" Then 		'validating english keyboard layout
						selectedWorkbookIndex = "0"
					End If
					
					result = excel_file.Assign_Workbook(selectedWorkbookIndex)	'assign doc
					cli.Print_Result(result)
					If result(0) <> -1 Then
						excel_file.activeWorkbook.Activate
						Exit Do
					End If
				Else
					Call Print_Message("Wrong input value!", "nok")
				End If
			Loop
		End If
		Attach_Workbook = result
	End Function

	Sub Select_Sheet()
		Dim result, selectedSheetIndex
		result = excel_file.Get_Excel_Application()
		If result(0) = -1 Then
			cli.Print_Result(result)
		Else
			Do While True
				'list open docs
				cli.Print_SubHeader("Sheets in workbook '" & excel_file.activeWorkbook.Name & "'")
				For i = 1 To excel_file.activeWorkbook.Sheets.Count
					cli.Print "(" & i & ") " & excel_file.activeWorkbook.Sheets(i).Name
				Next
				
				'select doc to work with
				cli.Print_Empty
				selectedSheetIndex = cli.Read_Input2("Selected sheet: ")
				
				'assigning doc
				If isNumeric(selectedSheetIndex) Then		'validating user input and 
					If selectedSheetIndex = "`" Then 		'validating english keyboard layout
						selectedSheetIndex = "0"
					End If
					
					result = excel_file.Select_Sheet(selectedSheetIndex)	'assign doc
					cli.Print_Result(result)
					If result(0) <> -1 Then
						Exit Do
					End If
				Else
					Call Print_Message("Wrong input value!", "nok")
				End If
			Loop
		End If
	End Sub

	Sub Clear_Range_Content()
		Dim result
		result = excel_file.Clear_Range_Content(cli.Read_Input2("Range: "))
		cli.Print_Result(result)
	End Sub

	Sub Clear_Range_Formatting()
		Dim result
		result = excel_file.Clear_Range_Formatting(cli.Read_Input2("Range: "))
		cli.Print_Result(result)
	End Sub

	Sub Clear_Range_Totally()
		Dim result
		result = excel_file.Clear_Range_Totally(cli.Read_Input2("Range: "))
		cli.Print_Result(result)
	End Sub
	
	Sub Clear_Sheet()
		Dim result
		result = excel_file.Clear_Sheet()
		cli.Print_Result(result)
	End Sub

	Sub Read_Range()
		Dim result
		result = excel_file.Read_Range(cli.Read_Input2("Range: "))
		cli.Print_Result(result)
	End Sub
	
	Sub Write_Range()
		Dim result
		result = excel_file.Read_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Value: "))
		cli.Print_Result(result)
	End Sub

	Sub Write_Range_With_Constant()
		Dim result
		result = excel_file.Read_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Value: "))
		cli.Print_Result(result)
	End Sub

	Sub Write_Formula_To_Range()
		Dim result
		result = excel_file.Read_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Formula: "))
		cli.Print_Result(result)
	End Sub
	
	Sub Set_Cell_Text_Color()
		Dim result
		result = excel_file.Set_Cell_Text_Color(cli.Read_Input2("Range: "), Array(cli.Read_Input2("[R]ed: "), cli.Read_Input2("[G]reen: "), cli.Read_Input2("[B]lue: ")))
		cli.Print_Result(result)
	End Sub
	
	Sub Set_Cell_Bg_Color()
		Dim result
		result = excel_file.Set_Cell_Bg_Color(cli.Read_Input2("Range: "), Array(cli.Read_Input2("[R]ed: "), cli.Read_Input2("[G]reen: "), cli.Read_Input2("[B]lue: ")))
		cli.Print_Result(result)
	End Sub

	Sub Get_Cell_Text_Color()
		Dim result
		result = excel_file.Get_Cell_Text_Color(cli.Read_Input2("Range: "))
		cli.Print("[R][G][B]: " & result(0) & ", " & result(1) & ", " & result(2))
	End Sub
	
	Sub Get_Cell_Bg_Color()
		Dim result
		result = excel_file.Get_Cell_Bg_Color(cli.Read_Input2("Range: "))
		cli.Print("[R][G][B]: " & result(0) & ", " & result(1) & ", " & result(2))
	End Sub
	
	
	Sub Search_First_Match_In_Range()
		Dim result
		result = excel_file.Search_First_Match_In_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Input: "), cli.Read_Input2("IsExactMatch: "), cli.Read_Input2("IsCaseSensitive: "))
		If result(0) <> -1 Then
			cli.Print("[" & result(1).Cells(1, 1).ColumnIndex & "," & result(1).Cells(1, 1).RowIndex & "] " & result(1).Cells(1, 1).Value)
		Else
			cli.Print("No cells found.")
		End If
	End Sub
	
	Sub Search_All_Matches_In_Range()
		Dim result
		result = excel_file.Search_All_Matches_In_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Input: "), cli.Read_Input2("IsExactMatch: "), cli.Read_Input2("IsCaseSensitive: "))
		If result(0) <> -1 Then
			For i = 1 To result(1).Rows.Count-1
				For j = 1 To result(1).Columns.Count-1
					cli.Print("[" & result(1).Cells(i, j).ColumnIndex & "," & result(1).Cells(i, j).RowIndex & "] " & result(1).Cells(i, j).Value)
				Next
			Next
		Else
			cli.Print("No cells found.")
		End If
	End Sub
	
	Sub Search_Index_Values_Of_First_Match_In_Range()
		Dim result
		result = excel_file.Search_Index_Values_Of_First_Match_In_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Input: "), cli.Read_Input2("IsExactMatch: "), cli.Read_Input2("IsCaseSensitive: "))
		If result(0) <> -1 Then
			cli.Print("[" & result(1).Cells(1, 1).ColumnIndex & "," & result(1).Cells(1, 1).RowIndex & "] " & result(1).Cells(1, 1).Value)
		Else
			cli.Print("No cells found.")
		End If
	End Sub
	
	Sub Search_Index_Values_Of_All_Matches_In_Range()
		Dim result
		result = excel_file.Search_Index_Values_Of_All_Matches_In_Range(cli.Read_Input2("Range: "), cli.Read_Input2("Input: "), cli.Read_Input2("IsExactMatch: "), cli.Read_Input2("IsCaseSensitive: "))
		If result(0) <> -1 Then
			For i = 1 To result(1).Rows.Count-1
				For j = 1 To result(1).Columns.Count-1
					cli.Print("[" & result(1).Cells(i, j).ColumnIndex & "," & result(1).Cells(i, j).RowIndex & "] " & result(1).Cells(i, j).Value)
				Next
			Next
		Else
			cli.Print("No cells found.")
		End If
	End Sub
	
	Sub Dump_Range()
		Dim result
		result = excel_file.Dump_Range(cli.Read_Input2("Range: "))
		cli.Print(result)
	End Sub
	
	Sub Dump_Range_Char_Code()
		Dim result
		result = excel_file.Dump_Range_Char_Code(cli.Read_Input2("Range: "))
		cli.Print(result)
	End Sub
End Class