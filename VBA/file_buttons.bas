Attribute VB_Name = "file_buttons"
Sub FileOpen_ppt()
' This function opens up the powerpoint template

 
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
        fullpath = .SelectedItems.Item(1)
    End With
    
    Range("ppt_template").Select
    ActiveCell.value = fullpath
 
End Sub
Sub SaveAs_ppt()
     
    Range("ppt_template").Select
    fullpath = ActiveCell.value

    Dim PowerPointApp As PowerPoint.Application
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    Dim myPresentation As PowerPoint.Presentation
    Set myPresentation = PowerPointApp.Presentations.Open(fullpath, WithWindow:=msoFalse)
      
    myPresentation.SaveAs Filename:=fullpath & "_test.pptx"
    
End Sub

Sub FileOpen_exceldata()
 
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
        fullpath = .SelectedItems.Item(1)
    End With
    
    Range("excel_data").Select
    ActiveCell.value = fullpath
 
End Sub

Sub set_dest_folder()

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
        dest_folder = .SelectedItems.Item(1)
    End With
    
    Range("dest_folder").Select
    ActiveCell.value = dest_folder

End Sub
