' Code related to ESET's VBA Dynamic Hook research
' For feedback or questions contact us at: github@eset.com
' https://github.com/eset/vba-dynamic-hook/
'
' This code is provided to the community under the two-clause BSD license as
' follows:
'
' Copyright (C) 2016 ESET
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice, this
' list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation
' and/or other materials provided with the distribution.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' Kacper Szurek <kacper.szurek@eset.com>
'
' Contain function wrappers and helpers

Public oblivion_fso As FileSystemObject
Public oblivion_log_object As TextStream
Public excel_object As Object
Public oblivion_excel_document As Object
Public oblivion_class_module As Object

Function oblivion_timestamp()
    Dim iNow
    Dim d(1 To 6)
    Dim i As Integer


    iNow = Now
    d(1) = Year(iNow)
    d(2) = Month(iNow)
    d(3) = Day(iNow)
    d(4) = Hour(iNow)
    d(5) = Minute(iNow)
    d(6) = Second(iNow)

    For i = 1 To 6
        If d(i) < 10 Then oblivion_timestamp = oblivion_timestamp & "0"
        oblivion_timestamp = oblivion_timestamp & d(i)
        If i = 3 Then oblivion_timestamp = oblivion_timestamp & " "
    Next i
End Function

Public Function oblivion_log(content As Variant)
	If TypeName(content) <> "String()" And TypeName(content) <> "Variant()" Then
		oblivion_log_object.Write StrConv(content, 0)
		oblivion_log_object.Write vbCrLf
	Else
	    oblivion_log_object.WriteLine "**Array**"
	End If
End Function


Public Sub log_return_from_string_function(name As Variant, content As Variant)
	If TypeName(content) <> "String()" And TypeName(content) <> "Variant()" Then
		oblivion_log_object.WriteLine name & " : " & StrConv(content, 0)
	End If
End Sub

Public Sub log_call_to_function(ParamArray a() As Variant)
    oblivion_log("External call: " & a(0) )

    Dim counter
    counter = 0
    For Each b In a
        if counter > 0 Then
            oblivion_log(vbTab & "Param " & counter)
            If TypeOf b Is Object Then
                oblivion_log (vbTab & vbTab & "object")
            Else:
                oblivion_log (vbTab & vbTab  & b)
            End If
        End if
        counter = counter + 1
    Next
End Sub

Function log_call_to_method(ParamArray d() As Variant)
Dim result As String
result = ""
For i = 0 To UBound(d)
    If TypeOf d(i) Is Object Then

    Else:
        If i > 0 Then
            result = result & ", " & d(i)
        Else:
            result = d(i)
        End If
    End If
Next i
oblivion_log result
End Function

Public Function Shell(PathName As Variant, Optional a As Variant) as Variant
	oblivion_log("Shell " & PathName)
	Shell = excel_object.Run("Shell_Builtin", PathName, a)
End Function

Public Function Mid(content As Variant, Start As Variant, Optional Length As Variant) As Variant
    Dim temp
    temp = excel_object.Run("Mid_Builtin", content, Start, Length)
    oblivion_log ("MID " & temp)
    Mid = temp
End Function

Public Function CreateObject(ObjectName As Variant) As Object
	oblivion_log("CreateObject " & ObjectName)
	Set CreateObject = excel_object.Run("CreateObject_Builtin", ObjectName)
End Function

Public Function GetObject(Optional a As Variant, Optional b as Variant) As Object
    If IsMissing(b) Then
        oblivion_log ("GetObject " & a)
        Set GetObject = excel_object.Run("GetObject_Builtin", a)
    Else
        If IsMissing(a) Then
            oblivion_log ("GetObject " & b)
            Set GetObject = excel_object.Run("GetObject_Builtin_2", b)
        Else
            oblivion_log ("GetObject " & a & " " & b)
            Set GetObject = excel_object.Run("GetObject_Builtin", a, b)
        End If
    End If
End Function

Public Function StrReverse(content as Variant) as Variant
	Dim temp
	temp = excel_object.Run("StrReverse_Builtin", content)
	oblivion_log("StrReverse " & temp)
	StrReverse = temp
End Function

Public Function Left(content As Variant, number_of_characters as Variant) as Variant
	Dim temp
	temp = excel_object.Run("Left_Builtin", content, number_of_characters)
	oblivion_log("Left " & temp)
	Left = temp
End Function

Public Function Environ(a as Variant) as Variant
    oblivion_log("Environ " & a)
    Environ = excel_object.Run("Environ_Builtin", a)
End Function

Public Function MsgBox(Prompt As Variant, Optional a As Variant, Optional b As Variant, Optional c As Variant, Optional d As Variant) as Variant
	oblivion_log("Messagebox " & Prompt)
	MsgBox = excel_object.Run("MsgBox_Builtin", Prompt)
End Function

Public Function InputBox(Prompt As Variant, Optional a As Variant, Optional b As Variant, Optional c As Variant, Optional d As Variant) as String
	oblivion_log("Inputbox " & Prompt)
	InputBox = word_object.Run("InputBox_Builtin", Prompt)
End Function

Public Sub oblivion_init()
	wrapper = "Public Function Mid_Builtin(content as Variant, Start As Variant, Optional Length As Variant) as Variant"
    wrapper = wrapper & vbLf & "    Mid_Builtin = Mid(content, Start, Length)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function CreateObject_Builtin(ObjectName As Variant) As Object"
    wrapper = wrapper & vbLf & "    Set CreateObject_Builtin = CreateObject(ObjectName)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function GetObject_Builtin(a As Variant, Optional b as Variant) As Object"
    wrapper = wrapper & vbLf & "    Set GetObject_Builtin = GetObject(a, b)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function GetObject_Builtin_2(b as Variant) As Object"
    wrapper = wrapper & vbLf & "    Set GetObject_Builtin_2 = GetObject(, b)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function Shell_Builtin(PathName As Variant, Optional a As Variant) as Variant"
    wrapper = wrapper & vbLf & "    Shell_Builtin = Shell(PathName)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function StrReverse_Builtin(content as Variant) as Variant"
    wrapper = wrapper & vbLf & "    StrReverse_Builtin = StrReverse(content)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function Left_Builtin(content As Variant, number_of_characters as Variant) as Variant"
    wrapper = wrapper & vbLf & "    Left_Builtin = Left(content, number_of_characters)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function Environ_Builtin(a as Variant) as Variant"
    wrapper = wrapper & vbLf & "    Environ_Builtin = Environ(a)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function MsgBox_Builtin(Prompt As Variant) as Variant"
    wrapper = wrapper & vbLf & "    MsgBox_Builtin = MsgBox(Prompt)"
    wrapper = wrapper & vbLf & "End Function"
    wrapper = wrapper & vbLf & "Public Function InputBox_Builtin(Prompt As Variant) as String"
    wrapper = wrapper & vbLf & "    InputBox_Builtin = InputBox(Prompt)"
    wrapper = wrapper & vbLf & "End Function"

    Set oblivion_fso = New FileSystemObject
    On Error Resume Next
    output_path = "<insert path here from oblivion>"
    If (Not oblivion_fso.FileExists(output_path & ".txt")) Then
        Set oblivion_log_object = oblivion_fso.CreateTextFile(output_path & ".txt", True)
    End If
    Set oblivion_log_object = oblivion_fso.OpenTextFile(output_path & ".txt", ForAppending, True)
    If Err.Number = 70 Then
        oblivion_log_object.Close
        Set oblivion_log_object = oblivion_fso.OpenTextFile(output_path & ".txt", ForAppending, True)
    End If
    Set excel_object = VBA.CreateObject("Excel.Application")
    Set oblivion_excel_document = excel_object.Workbooks.Add
    Set oblivion_class_module = oblivion_excel_document.VBProject.VBComponents.Add(1)
    oblivion_class_module.Name = "oblivion"
    oblivion_class_module.CodeModule.AddFromString wrapper
End Sub
