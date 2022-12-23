Set fso = CreateObject("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream(1)

'ExecuteGlobal fso.OpenTextFile("..\vbs_lib\word_utils.vbs").ReadAll()
'ExecuteGlobal fso.OpenTextFile("..\vbs_lib\cli_utils.vbs").ReadAll()
ExecuteGlobal fso.OpenTextFile("..\vbs_lib\excel_utils.vbs").ReadAll()

Dim word_utils, excel_utils, cli_utils, menu
Set word_utils = New WordUtils
Set excel_utils = New ExcelUtils
Set cli = New CliIo
Set menu = New CliMenu

'Choose Word or Excel file
Dim file_type
cli.Print_Header("General Office Tool")
cli.Print_Empty

Dim options
Set options = CreateObject("System.Collections.ArrayList")
options.Add("Word")
options.Add("Excel")
file_type = cli.Choose_Option("File type to use", options, "File type", "Wrong file type!", "File type '<option>' has been choosen!")


'Assign to open file and load functions with menu
Dim result
If file_type = 1 Then		'Word
	'Assign to open file
	result = word_utils.Attach_Document()
	If result(0) <> -1 Then
		'Load functions to menu	
		Call menu.AddCmd("Search Paragraph By Value", "word_utils.Search_Paragraph_By_Value")
		Call menu.AddCmd("Search Paragraph By Style", "word_utils.Search_Paragraph_By_Style")
		Call menu.AddCmd("Search All Paragraphs By Value", "word_utils.Search_All_Paragraphs_By_Value")
		Call menu.AddCmd("Search All Paragraphs By Style", "word_utils.Search_All_Paragraphs_By_Style")
		Call menu.AddCmd("Search Paragraph By Substr", "word_utils.Search_Paragraph_By_Substr")
		Call menu.AddCmd("Search All Paragraphs By Substr", "word_utils.Search_All_Paragraphs_By_Substr")
		Call menu.AddCmd("Search (universal)", "Search_Universal")
		Call menu.AddCmd("Dump Paragraph By Index", "word_utils.Dump_Paragraph_By_Index")
		Call menu.AddCmd("Dump Paragraph By Range", "word_utils.Dump_Paragraph_By_Range")
		Call menu.AddCmd("Dump Paragraph Char Codes", "word_utils.Dump_Paragraph_Char_Codes")
		Call menu.AddCmd("Dump Paragraph Char Codes By Range", "word_utils.Dump_Paragraph_Char_Codes_By_Range")
		Call menu.AddCmd("Print Statistics", "word_utils.Print_Statistics")
		Call menu.AddCmd("Set Heading Styles From Cli", "word_utils.Set_Heading_Styles_From_Cli")
		Call menu.AddCmd("Read All Headings", "word_utils.Read_All_Headings")
		Call menu.AddCmd("Dump All Headings By Group", "word_utils.Dump_All_Headings_By_Group")
		Call menu.AddCmd("Print Headings By Lvl", "word_utils.Print_Headings_By_Lvl")
		Call menu.AddCmd("Print Headings By Lvl And Parent", "word_utils.Print_Headings_By_Lvl_And_Parent")
		Call menu.AddCmd("Dump All Headings Raw", "word_utils.Dump_All_Headings_Raw")
		Call menu.AddCmd("Print Style Of Paragraph", "word_utils.Print_Style_Of_Paragraph")
		Call menu.Run
	End If
ElseIf file_type = 2 Then	'Excel
	'Assign to open file
	result = excel_utils.Attach_Workbook()
	If result(0) <> -1 Then
		'Load functions to menu	
		Call menu.AddCmd("Select Sheet", "excel_utils.Select_Sheet")
		Call menu.AddCmd("Clear Range Content", "excel_utils.Clear_Range_Content")
		Call menu.AddCmd("Clear Range Formatting", "excel_utils.Clear_Range_Formatting")
		Call menu.AddCmd("Clear Range Totally", "excel_utils.Clear_Range_Totally")
		Call menu.AddCmd("Clear Sheet", "excel_utils.Clear_Sheet")
		Call menu.AddCmd("Read Range", "excel_utils.Read_Range")
		Call menu.AddCmd("Write Range", "excel_utils.Write_Range")
		Call menu.AddCmd("Write Range_With_Constant", "excel_utils.Write_Range_With_Constant")
		Call menu.AddCmd("Write Formula_To_Range", "excel_utils.Write_Formula_To_Range")
		Call menu.AddCmd("Set Cell Text Color", "excel_utils.Set_Cell_Text_Color")
		Call menu.AddCmd("Set Cell Bg Color", "excel_utils.Set_Cell_Bg_Color")
		Call menu.AddCmd("Get Cell Text Color", "excel_utils.Get_Cell_Text_Color")
		Call menu.AddCmd("Get Cell Bg Color", "excel_utils.Get_Cell_Bg_Color")
		Call menu.AddCmd("Search First_Match In_Range", "excel_utils.Search_First_Match_In_Range")
		Call menu.AddCmd("Search All_Matches In_Range", "excel_utils.Search_All_Matches_In_Range")
		Call menu.AddCmd("Search Index Values Of First Match In Range", "excel_utils.Search_Index_Values_Of_First_Match_In_Range")
		Call menu.AddCmd("Search Index Values Of All Matches In Range", "excel_utils.Search_Index_Values_Of_All_Matches_In_Range")
		Call menu.AddCmd("Dump Range", "excel_utils.Dump_Range")
		Call menu.AddCmd("Dump Range Char Code", "excel_utils.Dump_Range_Char_Code")
		Call menu.Run
	End If
End If

Sub Search_Universal()
	'later it should be defined as an universal function with dinamic count of options to choose in the cli_utils file
	Dim type1, type2, type3
	type1 = ""
	type2 = ""
	type3 = ""
	
	'First match or all matches?
	Do While True
		type1 = UCase(cli.Read_Input2("[F]irst match (default) or [A]ll matches: "))
		If type1 = "F" Or type1 = "A" Or type1 = "" Then
			If type1 = "" Then
				type1 = "F"
			End If
			Exit Do
		Else
			cli.Print_Nok("Wrong input!")
		End If
	Loop
	
	'By value or by style?
	Do While True
		type2 = UCase(cli.Read_Input2("By [V]alue (default) or By [S]tyle: "))
		If type2 = "V" Or type2 = "S" Or type2 = "" Then
			If type2 = "" Then
				type2 = "V"
			End If
			Exit Do
		Else
			cli.Print_Nok("Wrong input!")
		End If
	Loop
	
	'Exact match or contains string? (Just if by value has been choosen previously)
	If type2 = "V" Then
		Do While True
			type3 = UCase(cli.Read_Input2("[E]xact match (default) or [C]ontains substring: "))
			If type3 = "E" Or type3 = "C" Or type3 = "" Then
				If type3 = "" Then
					type3 = "E"
				End If
				Exit Do
			Else
				cli.Print_Nok("Wrong input!")
			End If
		Loop
	End If
	
	'Call the appropriate search function
	If type1 = "F" Then
		If type2 = "V" Then
			If type3 = "E" Then
				word_utils.Search_Paragraph_By_Value
			ElseIf type3 = "C" Then
				word_utils.Search_Paragraph_By_Substr
			End If
		ElseIf type2 = "S" Then
			word_utils.Search_Paragraph_By_Style
		End If
	ElseIf type1 = "A" Then
		If type2 = "V" Then
			If type3 = "E" Then
				word_utils.Search_All_Paragraphs_By_Value
			ElseIf type3 = "C" Then
				word_utils.Search_All_Paragraphs_By_Substr
			End If
		ElseIf type2 = "S" Then
			word_utils.Search_All_Paragraphs_By_Style
		End If
	End If
End Sub


