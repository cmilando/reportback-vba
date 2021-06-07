VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formatting_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Chad Milando
' Copyright : Copyright 2020, 2021, Trustees of Boston University.
'             All rights reserved.
' License   : This software is provided under the terms of the
'             SOFTWARE EVALUATION LICENSE AGREEMENT as detailed in
'             the "license.txt" file in source directory.
'---------------------------------------------------------------------------------------
'@Folder("VBAProject")
Public str_txt As String
Public str_type As String
Public start As Integer
Public length As Integer

Sub set_values(fmt_i As Integer, str_txt_i As Variant, str_type_i As Variant, _
            start_i As Variant, length_i As Variant)
    
    ' type
    If Not value_in_set(VarType(str_type_i), Array(8)) Then
        MsgBox "ERROR: 'type' should be string - see row " & fmt_i & _
               " with value " & str_type_i
        End
    End If
    
    ' type
    If Not value_in_set(str_type_i, Array("superscript", "subscript")) Then
        MsgBox "ERROR: 'type' should one of 'superscript' or 'subscript' " & _
            " - see row " & fmt_i & " with value " & str_type_i
        End
    End If
       
    ' start
    If Not value_in_set(VarType(start_i), Array(2, 3, 4, 5)) Then
        MsgBox "ERROR: 'start' should be numeric - see row " & fmt_i & _
               " with value " & start_i
        End
    End If
    
    If start_i > Len(str_txt_i) Or start_i < 0 Then
        MsgBox "ERROR: invalid start integer, must be 1 <= x <= len(text)" & _
               ", see row " & fmt_i
        End
    End If
               
    ' length
    If Not value_in_set(VarType(length_i), Array(2, 3, 4, 5)) Then
        MsgBox "ERROR: 'length' should be numeric - see row " & fmt_i & _
               " with value " & length_i
        End
    End If
    
    If length_i > Len(str_txt_i) Or length_i < 0 Or _
       length_i + start_i - 1 > Len(str_txt_i) Then
        MsgBox "ERROR: invalid length integer, must be 1 <= x <= len(text) - start" & _
               ", see row " & fmt_i
        End
    End If

    str_txt = str_txt_i
    str_type = str_type_i
    start = CInt(start_i)
    length = CInt(length_i)

End Sub

' ! ===========================================================================
Function value_in_set(strIn As Variant, arrList As Variant)
    
    value_in_set = Not IsError(Application.Match(strIn, arrList, 0))
        
End Function


