Attribute VB_Name = "file_buttons"
'---------------------------------------------------------------------------------------
' Author    : Chad Milando
' Copyright : Copyright 2020, 2021, Trustees of Boston University.
'             All rights reserved.
' License   : This software is provided under the terms of the
'             SOFTWARE EVALUATION LICENSE AGREEMENT as detailed in
'             the "license.txt" file in source directory.
'---------------------------------------------------------------------------------------
Option Explicit
'! ============================================================================
Sub RangeExists(WhatSheet As String, WhatRange As String, WhatCell As String)
    Dim test As Range
    Dim RangeExists As Boolean
    
    On Error Resume Next
    Set test = ActiveWorkbook.Sheets(WhatSheet).Range(WhatRange)
    RangeExists = Err.Number = 0
    On Error GoTo 0
    
    If RangeExists <> True Then
        MsgBox ("ERROR: Named range '" & WhatRange & _
                 "' does not exist (should be Cell " & _
                 WhatCell & ")")
        End
    End If
    
End Sub

Sub validate_option(WhatRange As String, arrList As Variant)
    Dim test As Boolean
    Dim strIn As String
    
    strIn = LCase(Range(WhatRange).value)
    
    test = Not (IsError(Application.Match(strIn, arrList, 0)))
    
    If test <> True Then
        MsgBox ("ERROR: Named range '" & WhatRange & _
                 "' value (" & strIn & ") not in expected options (" & _
                 Join(arrList, ", ") & ")")
        End
    End If
    
End Sub

'! ============================================================================
Sub FileOpen_ppt()

' This function opens up the template
    Dim fullpath As String
    
    Call RangeExists("Run", "template", "D15")
    Call RangeExists("Run", "left_char", "C18")
    Call RangeExists("Run", "right_char", "C19")
    
    Call validate_option("left_char", Array("{", "{#"))
    Call validate_option("right_char", Array("}", "#}"))
    
    'Display a Dialog Box that allows to select a single file.
    'The path for the file picked will be stored in fullpath variable
    With Application.FileDialog(msoFileDialogFilePicker)
          'Makes sure the user can select only one file
          .AllowMultiSelect = False
          'Filter to just the following types of files to narrow down selection options
          .Filters.Add "Powerpoint Files", "*.ppt; *.pptx; *.pptm", 1
          'Show the dialog box
          .Show
          
          'Store in fullpath variable
          If .SelectedItems.Count = 0 Then
              fullpath = Range("template").value
          Else
              fullpath = .SelectedItems.Item(1)
          End If
      End With
      
      Range("template").Select
      ActiveCell.value = fullpath
 
End Sub

'! ============================================================================
Sub FileOpen_exceldata()

 Dim fullpath As String
 
 Call RangeExists("Run", "excel_data", "D27")
 Call RangeExists("Run", "use_image_data", "C32")
 Call RangeExists("Run", "use_formatting_data", "C33")
 
 Call validate_option("use_image_data", Array("yes", "no"))
 Call validate_option("use_formatting_data", Array("yes", "no"))
 
'Display a Dialog Box that allows to select a single file.
'The path for the file picked will be stored in fullpath variable
  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        If .SelectedItems.Count = 0 Then
            fullpath = Range("excel_data").value
        Else
            fullpath = .SelectedItems.Item(1)
        End If
    End With
    
    Range("excel_data").Select
    ActiveCell.value = fullpath
 
End Sub

'! ============================================================================
Sub set_dest_folder()

    Dim dest_folder As String
    Dim diaFolder As FileDialog
    Dim selected As Long
    
    Call RangeExists("Run", "dest_folder", "D37")
    Call RangeExists("Run", "output_as", "C40")
    Call RangeExists("Run", "output_suffix", "C41")
    
    Call validate_option("output_as", Array("ppt", "pdf"))
    Call validate_option("output_suffix", Array("date", "none"))
    
    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show

    With Application.FileDialog(msoFileDialogFolderPicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Show the dialog box
        '.Show
        
        'Store in fullpath variable
        If .SelectedItems.Count = 0 Then
            dest_folder = Range("dest_folder").value
        Else
            dest_folder = .SelectedItems.Item(1)
        End If
        
    End With
    
    Range("dest_folder").Select
    ActiveCell.value = dest_folder

End Sub
