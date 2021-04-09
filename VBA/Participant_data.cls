VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Participant_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ! ===========================================================================
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
' Member variables - private means they are not available
' These are variables to pass between subs that you eventually want
' to make Property Get for
'
Private wb As Excel.Workbook         ' workbook with participant data
Private ws As Excel.Worksheet        ' individual worksheets

Private n_pp As Integer              ' number of people in excel_data,
Private pp_ids() As Variant          ' person ids from the first column
Private pp_data() As Variant         ' all of person data
Private tmp As Variant               '

Private n_vars As Integer            ' number of variables in excel_data
Private var_names() As String        ' list of variable names

Private use_image_data As Boolean    '
Private n_images_total As Integer    '
Private image_coll As New Collection '

Private use_formatting As Boolean    '
Private n_fmt_unique As Integer      '
Private fmt_db() As Variant          '
Private fmt_unique() As Variant      '

Private Logger() As Variant          ' re-dimable logger
Private logger_i As Integer          ' iterate through logger
Private logprint_i As Integer        ' iterate through logger

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
   
    Set wb = GetObject(Range("excel_data").value)
        
    ' Clear the error log page, and set up the logger
    Const logger_size As Integer = 1000 ' Logger size
    ReDim Logger(logger_size)
    logger_i = 1
    logprint_i = 1
    Call SheetExists(ActiveWorkbook, "ErrorLog")
    Sheets("ErrorLog").Cells.Clear
    
End Sub
' ! ===========================================================================
' Sub: Check if a sheet exists
Sub SheetExists(wb As Workbook, WhatSheet As String)

    Dim test As Worksheet
    Dim SheetExists As Boolean
    
    On Error Resume Next
    Set test = wb.Worksheets(WhatSheet)
    SheetExists = Err.Number = 0
    On Error GoTo 0
    
    If SheetExists <> True Then
        MsgBox ("ERROR: Sheet '" & WhatSheet & "' does not exist - " & _
        "check spelling.")
        End
    End If
    
End Sub

' ! ===========================================================================
' Sub: fill_vars()
' Purpose:
'
' Validation to check:
' - is there actually a sheet called text_data
' - does it start on A1
' - also must be continuous and have no breaks
Public Sub fill_vars()
    
    ' Check to make sure that various sheets exist if used
    Call SheetExists(wb, "text_data")
    
    Set ws = wb.Worksheets("text_data")
      
    Dim pp_i As Integer             ' person counter
    Dim var_i As Integer            ' variable counter
    
    'fill pp_ids, var_data and vars_to_find
    n_pp = 0
    n_vars = 0
    Do While ws.Range("A2").Offset(n_pp, 0).value <> ""
        n_pp = n_pp + 1
    Loop
    Do While ws.Range("B1").Offset(0, n_vars).value <> ""
        n_vars = n_vars + 1
    Loop
    If n_pp = 0 Or n_vars = 0 Then
        MsgBox "ERROR: Blank initial row or column on sheet 'text_data'"
        End
    End If
    
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
                tmp = ws.Range("A1").Offset(pp_i, 0).value
                
                ' make sure this pp_id doesn't already exist
                If value_in_set(tmp, pp_ids) Then
                    MsgBox "ERROR: PPids cannot be doubled - see row " & pp_i
                    End
                End If
                
                pp_ids(pp_i) = tmp
                
            End If
            
            tmp = ws.Range("A1").Offset(pp_i, var_i).value
            
            ' Value validation
            If value_in_set(tmp, Array("", "NA", "<NA>")) Then
                Logger(logger_i) = "Check value of Person: " & pp_ids(pp_i) & _
                                   " and Var: " & var_names(var_i) & ". " & _
                                   "Recorded value is in c('', 'NA', '<NA>')."
                                   
                logger_i = logger_i + 1
            End If
            
            pp_data(pp_i, var_i) = tmp
            
        Next pp_i
        
    Next var_i
    
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Column names = " & Join(var_names, ", ")
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
    Progress_bar.HelpText.Caption = "Person ids  = " & Join(pp_ids, ", ")
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
End Sub

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
    
    ' Check to make sure that various sheets exist if used
    Call SheetExists(wb, "image_data")
    
    Set ws = wb.Worksheets("image_data")
    
    Dim image_i As Integer          '
    Dim image_var_i As Integer
    Dim n_image_cols As Integer
    Dim sub_image_coll As New Collection
    Dim key As String
    Dim image As New Image_data
    
    Dim pp As Variant
    Dim slide As Variant
    Dim path As Variant
    Dim top  As Variant
    Dim left  As Variant
    Dim height As Variant
    Dim width As Variant
    
    ' different people might have different number of files
    n_images_total = 0
    n_image_cols = 0
    Do While ws.Range("A2").Offset(n_images_total, 0).value <> ""
        n_images_total = n_images_total + 1
    Loop
    Do While ws.Range("A1").Offset(0, n_image_cols).value <> ""
        n_image_cols = n_image_cols + 1
    Loop
    If n_images_total = 0 Or n_image_cols <> 7 Then
        MsgBox "ERROR: Blank initial row or necessary column on 'image_data'"
        End
    End If
    
    'pp id, slide#, path, top, left, height, width = 7
    ReDim image_db(n_images_total, 7)
        
    For image_i = 1 To n_images_total
    
        Call check_cancel
        
        ' ID must be in pp_ids
        pp = ws.Range("A1").Offset(image_i, 0).value
        If Not value_in_set(pp, pp_ids) Then
            MsgBox "ERROR: Image ID not in text_data ID - see row " & image_i & _
                   " with value " & pp
            End
        End If
        
        ' Create a new image object
        slide = ws.Range("A1").Offset(image_i, 1).value
        path = ws.Range("A1").Offset(image_i, 2).value
        top = ws.Range("A1").Offset(image_i, 3).value
        left = ws.Range("A1").Offset(image_i, 4).value
        height = ws.Range("A1").Offset(image_i, 5).value
        width = ws.Range("A1").Offset(image_i, 6).value
        
        Set image = New Image_data
        Call image.set_values(pp, slide, path, top, left, width, height)

        ' Insert it into the collection
        key = pp & "_" & slide
        If Not Exists(image_coll, key) Then
        
            Set sub_image_coll = New Collection
            
            sub_image_coll.Add image
            
            image_coll.Add sub_image_coll, key
            
        Else
        
            image_coll(key).Add image
        
        End If
        
    Next image_i
    
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Total # of images = " & n_images_total
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Call check_cancel
    
End Sub

' -------------------------------------
' get image data
Property Get n_images() As Integer
    n_images = n_images_total
End Property

Property Get images() As Collection
    Set images = image_coll
End Property

' ! ===========================================================================
' Sub: fill_formatting_chars()
' Purpose:
'
' Validation to check:
' - is there actually a sheet called formatting
' - does it start on A1
' - are the formatting instructions valid (eg in a list of approved answers)
Public Sub fill_formatting_chars()
    
    ' Check to make sure that various sheets exist if used
    Call SheetExists(wb, "formatting")
    
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
        
        ' Validation
        ' Check that the n chars is good
        ' check that the value given can actually be in (is in the set)
        
    Next fmt_i
    
    fmt_unique(n_fmt_unique, 3) = fmt_i - 1
    
    'helpful message boxes before moving on
    Progress_bar.HelpText.Caption = "Total # of formatting = " & n_fmt_total
    Application.Wait (Now + TimeValue("0:00:01"))
    Call check_cancel
    
End Sub

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

' ! ===========================================================================
Sub print_logger()

    If logger_i > 1 Then
        
        For logprint_i = 1 To (logger_i - 1)
            
            Sheets("ErrorLog").Cells(1 + logprint_i, 2).value = Logger(logprint_i)
            
        Next logprint_i
            
        MsgBox ("CHECK ERROR LOG")
        
    End If

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
' Purpose:
Function value_in_set(strIn As Variant, arrList As Variant)
    
    value_in_set = Not IsError(Application.Match(strIn, arrList, 0))
        
End Function
' ! ===========================================================================
' Purpose:
Function type_is(strIn As Variant, expected_type As Integer)

    If expected_type = VarType(strIn) Then
        type_is = True
    Else
        type_is = False
    End If
    
End Function
' ! ===========================================================================
Function Exists(coll As Collection, key As String) As Boolean

    On Error GoTo EH

    IsObject (coll.Item(key))
    
    Exists = True
EH:
End Function
