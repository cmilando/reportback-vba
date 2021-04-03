VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Participant_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: CWM
' Class Module: participant_data
'
' # from https://excelmacromastery.com/vba-class-modules/
' Methods – functions/subs.
' Member variables – variables.
' Properties– types of functions/subs that behave like variables.
' Events – subs that are triggered by an event.

' example of get vs let
' https://www.reddit.com/r/vba/comments/f2qzti/what_exactly_is_the_purpose_of_property_letget/

Option Explicit

' ! ===========================================================================
' Member variables
' These are variables to pass between subs that you eventually want
' to make Property Get for
Private wb As Excel.Workbook        ' workbook with participant data
Private ws As Excel.Worksheet       ' individual worksheets

Private n_pp As Integer             ' number of people in excel_data,
Private pp_ids() As Variant         ' person ids from the first column
Private pp_data() As Variant        ' all of person data

Private n_vars As Integer           ' number of variables in excel_data
Private var_names() As String       ' list of variable names

Private n_images_total As Integer   '
Private image_db() As Variant       '

Private n_fmt_unique As Integer     '
Private fmt_db() As Variant         '
Private fmt_unique() As Variant     '


' ! ===========================================================================
' Sub: Class_Initialize()
' Purpose:
'
' Validation to check:
' - the file in excel_data is valid
' - is there the sheets text_data, image_data, formatting
' - what if one of these sheets is missing?
' - what if someone just has images, or just has data?
' - assume they always have data right? the other two are optional
'   if they dont want to print data, then they have to just make a dummy
'   column on text data that doesn't show up anywhere in the ppt
Private Sub Class_Initialize()

    'excel_data file must be closed
    
    'also must be continuous and have no breaks
    Progress_bar.HelpText.Caption = "Open excel data workbook"
    Set wb = GetObject(Range("excel_data").value)

End Sub

' ! ===========================================================================
' Sub: fill_vars()
' Purpose:
'
' Validation to check:
' - is there actually a sheet called text_data
' - does it start on A1
Public Sub fill_vars()
    
    Set ws = wb.Worksheets("text_data")
      
    Dim pp_i As Integer             ' person counter
    Dim var_i As Integer            ' variable counter
    
    'fill pp_ids, var_data and vars_to_find
    ' <<<< FIX THIS, WONT WORK IF # DATA COLUMNS < 2
    n_pp = ws.Range("A2").End(xlDown).Row - 1
    n_vars = ws.Range("B1").End(xlToRight).Column - 1
    ReDim var_names(n_vars)
    ReDim pp_data(n_pp, n_vars)
    ReDim pp_ids(n_pp)
        
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Number of People = " & n_pp & _
                                    ". Number of variables = " & n_vars
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
    ' Now loop through and fill with data
    For var_i = 1 To n_vars
    
        var_names(var_i) = ws.Range("A1").Offset(0, var_i).value
        
        For pp_i = 1 To n_pp
            
            Call check_cancel
            
            If var_i = 1 Then
                pp_ids(pp_i) = ws.Range("A1").Offset(pp_i, 0).value
            End If
            pp_data(pp_i, var_i) = ws.Range("A1").Offset(pp_i, var_i).value
            
        Next pp_i
        
    Next var_i
    
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Column names = " & Join(var_names, vbCrLf)
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
    Progress_bar.HelpText.Caption = "Person ids  = " & Join(pp_ids, vbCrLf)
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
End Sub

' ! ===========================================================================
' Sub: fill_image_data()
' Purpose:
'
' Validation to check:
' - is there actually a sheet called image_data
' - does it start on A1
' - what if the slides are out of order
' - or if person is non-consecutive
Public Sub fill_image_data()

    Set ws = wb.Worksheets("image_data")
    
    Dim image_i As Integer          '
    Dim image_var_i As Integer      '
    
    ' different people might have different number of files
    n_images_total = ws.Range("A2").End(xlDown).Row - 1
    'pp id, slide#, path, top, left, height, width = 7
    ReDim image_db(n_images_total, 7)
        
    For image_i = 1 To n_images_total
    
        Call check_cancel
    
        For image_var_i = 0 To 6
            image_db(image_i, image_var_i + 1) = _
                           ws.Range("A1").Offset(image_i, image_var_i).value
        Next image_var_i
        
    Next image_i
    
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Total # of images = " & n_images_total
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
End Sub

' ! ===========================================================================
' Sub: fill_formatting_chars()
' Purpose:
'
' Validation to check:
' - is there actually a sheet called formatting
' - does it start on A1
Public Sub fill_formatting_chars()

    Set ws = wb.Worksheets("formatting")
    
    Dim n_fmt_total As Integer      '
    Dim fmt_i As Integer            '
    Dim fmt_var_i As Integer        '
    Dim unique_i As Integer         '
    
    n_fmt_total = ws.Range("A2").End(xlDown).Row - 1
    ReDim fmt_unique(n_fmt_total, 3) 'fmt, start, stop
    ReDim fmt_db(n_fmt_total, 4) 'text, type, start ,len

    n_fmt_unique = 0
    unique_i = 1
    
    For fmt_i = 1 To n_fmt_total
        
        Call check_cancel
        
        For fmt_var_i = 0 To 3
            fmt_db(fmt_i, fmt_var_i + 1) = _
                            ws.Range("A1").Offset(fmt_i, fmt_var_i).value
        Next fmt_var_i
        
        If fmt_i = 1 Then
            n_fmt_unique = 1
            fmt_unique(n_fmt_unique, 1) = fmt_db(fmt_i, 1) 'name
            fmt_unique(n_fmt_unique, 2) = fmt_i
            
        Else
            If fmt_db(fmt_i, 1) <> fmt_db(fmt_i - 1, 1) Then
                fmt_unique(n_fmt_unique, 3) = fmt_i - 1
                n_fmt_unique = n_fmt_unique + 1
                fmt_unique(n_fmt_unique, 1) = fmt_db(fmt_i, 1)
                fmt_unique(n_fmt_unique, 2) = fmt_i
            End If
        End If
        
    Next fmt_i
    
    fmt_unique(n_fmt_unique, 3) = fmt_i - 1
    
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Total # of formatting = " & n_fmt_total
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
End Sub

' ! ===========================================================================
' Sub: check_cancel()
' Purpose:
Sub check_cancel()

    DoEvents
    If stopit = True Then
        MsgBox ("Execution stopped")
        End
    End If

End Sub

' ! ===========================================================================
' Public properties

' -------------------------------------------
' get vars_to_find
Property Get keys() As Variant
    keys = var_names
End Property

' get n_keys
Property Get n_keys() As Integer
    n_keys = n_vars
End Property

' -------------------------------------------
' get pp_ids
Property Get person_ids() As Variant
    person_ids = pp_ids
End Property

' get n_pp
Property Get n_persons() As Integer
    n_persons = n_pp
End Property

' get var_data
Property Get var_data() As Variant
    var_data = pp_data
End Property

' --------------------------------------------
' get image data
Property Get n_images() As Integer
    n_images = n_images_total
End Property

Property Get images() As Variant
    images = image_db
End Property

' -------------------------------------------
' get formatted character data
Property Get fmt_u() As Variant
    fmt_u = fmt_unique
End Property

Property Get fmt_data() As Variant
    fmt_data = fmt_db
End Property

Property Get n_fmt() As Variant
    n_fmt = n_fmt_unique
End Property



