Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetAbsolutePathName(".")  & "\"
file_path = fso.GetFile(WScript.Arguments.Item(0))
file_name = fso.GetFileName(file_path)
file_ext = Split(file_name, ".")(1)
instrumented_macro_path = fso.GetFile(WScript.Arguments.Item(1))
clean_file_path = CurrentDirectory & "resources\clean_files\clean." + file_ext
output_file_path = Replace(file_path, "." + file_ext, "") & "_output." + file_ext
injection_flag = False

' This creates the word application and takes the macro content prepared with the parser
If file_ext = "xls" OR file_ext = "xlsm" Then
    vhook_module_path = CurrentDirectory & "bin\class_excel.vba"
    Set office_app = CreateObject("Excel.Application")
    office_app.Visible = False
    office_app.DisplayAlerts = False
    Set office_doc = office_app.Workbooks.Open(clean_file_path)
    main_macro_name = "ThisWorkbook.cls"
ElseIf file_ext = "doc" OR file_ext = "docm" Then
    vhook_module_path = CurrentDirectory & "bin\class_word.vba"
    Set office_app = CreateObject("Word.Application")
    office_app.Visible = False
    office_app.DisplayAlerts = False
    Set office_doc = office_app.Documents.Open(clean_file_path)
    main_macro_name = "ThisDocument.cls"
End If

Set macro_file = fso.OpenTextFile(instrumented_macro_path, 1)
str_macro_code = macro_file.ReadAll
macro_file.Close

'if ----- is in macro, there are more than one macro inside the file
If Instr(str_macro_code, "-------------------------------------------------------------------------------") <> 0 Then
	arr_macro_code = Split(str_macro_code, "-------------------------------------------------------------------------------", -1, 1)
	For Each macro_str In arr_macro_code
        module_name = Left(macro_str, Instr(macro_str, "~~") - 1)
        module_name = Replace(module_name, "VBA MACRO ", "")
        If module_name = main_macro_name Then
            ' Injects the macro content
            macro_content = Trim(Split(macro_str, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", -1, 1)(1))
            Set VarComp = office_doc.VBProject.VBComponents(1)
            VarComp.CodeModule.AddFromString macro_content
        Else
            macro_content = Trim(Split(macro_str, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", -1, 1)(1))
            If Len(Replace(macro_content, vbCrLf, "")) > 0 Then
                module_name = Replace(module_name, vbCrLf, "")
                module_name = Left(module_name, Instr(module_name, ".") - 1)
                If file_ext = "xls" OR file_ext = "xlsm" Then
                    On Error Resume Next
                        sheet_list = office_doc.WorkSheets
                        If Err.Number <> 0 Then
                            Err.Clear
                        Else
                            For Each sheet In sheet_list
                                If sheet.Name = module_name Then
                                    Set macro_module = office_doc.VBProject.VBComponents(module_name)
                                    macro_module.CodeModule.AddFromString macro_content
                                    injection_flag = True
                                End If
                            Next
                        End If
                    If injection_flag = False Then
                        Set macro_module = office_doc.VBProject.VBComponents.Add(1)
                        macro_module.Name = module_name
                        macro_module.CodeModule.AddFromString macro_content
                    End If
                    injection_flag = False
                Else
                    if module_name <> "VBA_P-code" Then
                        Set macro_module = office_doc.VBProject.VBComponents.Add(1)
                        macro_module.Name = module_name
                        macro_module.CodeModule.AddFromString macro_content
                    End If
                End If
            End If
        End If
	Next
Else
	' Injects the macro content
	Set VarComp = office_doc.VBProject.VBComponents(1)
	macro_content = Split(str_macro_code, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", -1, 1)(1)
	VarComp.CodeModule.AddFromString macro_content
End If

' Add the vhook logging routines
On Error Resume Next
    office_doc.VBProject.References.AddFromGUID "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
    If Err.Number <> 0 Then
        Err.Clear
    End If
Dim vhook_module_file, vhook_module_content
Set vhook_module_file = fso.OpenTextFile(vhook_module_path)
vhook_module_content = vhook_module_file.ReadAll
vhook_module_file.Close
Set macro_module = office_doc.VBProject.VBComponents.Add(1)
macro_module.Name = "vhook"
macro_module.CodeModule.AddFromString vhook_module_content

office_doc.SaveAs output_file_path
office_doc.Close
On Error Resume Next
    office_app.Quit
    If Err.Number <> 0 Then
        Err.Clear
    End If

Do
	On Error Resume Next
	If file_ext = "xls" OR file_ext = "xlsm" Then
        Set objWord = GetObject(, "Excel.Application")
    ElseIf file_ext = "doc" OR file_ext = "docm" Then
        Set objWord = GetObject(, "Word.Application")
    End If
	If Not objWord Is Nothing Then
		objWord.Quit
		Set objWord = Nothing
	End If
Loop Until objWord Is Nothing

If file_ext = "xls" OR file_ext = "xlsm" Then
    Set office_app = CreateObject("Excel.Application")
    office_app.Visible = True
    office_app.DisplayAlerts = False
    Set office_doc = office_app.Workbooks.Open(output_file_path)
ElseIf file_ext = "doc" OR file_ext = "docm" Then
    Set word_app = CreateObject("Word.Application")
    word_app.Visible = True
    word_app.DisplayAlerts = False
    Set word_doc = word_app.Documents.Open(output_file_path)
End IF

