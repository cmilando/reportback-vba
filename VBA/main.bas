Attribute VB_Name = "main"
'@Folder "VBAProject"
Option Explicit
Public stopit As Boolean
Sub main()

    UserForm1.Show
    
End Sub

Sub progress(pctCompl As Single)

    UserForm1.Text.Caption = pctCompl & "% Completed"
    UserForm1.Bar.Width = pctCompl * 2
    
    DoEvents

End Sub

Sub validation()

    ' Validation


End Sub


Sub find_and_replace()
    
    ' never add to more than 3 parameters to a method
    ' never indent more than 3 times in a method
    ' try and keep a single function under 100 lines --> single responsiblity principle
    ' readme data function that gives back data
    ' make functions for each of the sub-routines
    ' can i define classes
    ' >> load data, and give you back a class that represented
    ' >> can i define a class for a person? that I can pass around
    ' >> can I define class methods for each person
    ' >> make a class for slides, persons,
    ' >> public class my powerpoint
    ' >> on it i can put public properties
    
    ' put this above the sub find and replace
    ' Public Class Customer
    '   Public Property AccountNumber As Integer
    ' End Class
    ' collection object?
    ' pull the data this easily accessible
    
    ' complexity is unavoidable
    ' manage the over-head
    
    ' should be able to do dictionaries
    ' should be able to do classes
    ' should be able to make a good state thing
    ' this class structure
    ' put the class structure as a field on the class
    ' the fill thing returns true or false
    ' all have different loggers that do the same thing
    ' this is why its good to have a logging classes
    ' STATIC logger is the same no matter what person you use
    ' it can be a field of the class so

    
    'need to add the powerpoint library to tools > references
    stopit = False
    Application.ScreenUpdating = False
    UserForm1.HelpText.Caption = "Starting"
    
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 0: CLEAR ERROR LOG, VALIDATION CHECKS
    ' ---------------------------------------------------------------------------------------------------------
    
    ' Validation
    Call validation
    
    ' Clear the error log page, and set up the logger
    Dim logger() As Variant                     ' re-dimable logger
    Dim logger_i, logprint_i As Integer         ' iterate through logger
    Const logger_size As Integer = 1000 ' Logger size
    ReDim logger(logger_size)
    logger_i = 1
    logprint_i = 1
    Sheets("ErrorLog").Cells.Clear

    
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 1A: READ IN VARIABLES FROM TEXT DATA
    ' ---------------------------------------------------------------------------------------------------------
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    
    'excel_data file must be closed
    'also must be continuous and have no breaks
    UserForm1.HelpText.Caption = "Open excel data workbook"
    Set wb = GetObject(Range("excel_data").Value)
    Set ws = wb.Worksheets("text_data")
      
    Dim n_pp, pp_i As Integer       'number of people in excel_data, person counter
    Dim pp_ids() As Variant         'person ids from the first column
    Dim n_vars, var_i As Integer    'number of variables in excel_data, variable counter

    Dim vars_to_find() As String    'list of variable names
    Dim var_data() As Variant       'all of variable data
    
    'fill pp_ids, var_data and vars_to_find
    n_pp = ws.Range("A2").End(xlDown).Row - 1
    n_vars = ws.Range("B1").End(xlToRight).Column - 1  ' <<<< FIX THIS, WONT WORK IF # DATA COLUMNS < 2
    UserForm1.HelpText.Caption = "Number of People = " & n_pp & ". Number of variables = " & n_vars
    Application.Wait (Now + TimeValue("0:00:01"))
    ReDim vars_to_find(n_vars)
    ReDim var_data(n_pp, n_vars)
    ReDim pp_ids(n_pp)
    
    For var_i = 1 To n_vars
        vars_to_find(var_i) = ws.Range("A1").Offset(0, var_i).Value
        For pp_i = 1 To n_pp
            
            DoEvents
            If stopit = True Then
                UserForm1.HelpText.Caption = "execution stopped. press 'x' to close"
                Exit Sub
            End If
                
            If var_i = 1 Then
                pp_ids(pp_i) = ws.Range("A1").Offset(pp_i, 0).Value
            End If
            var_data(pp_i, var_i) = ws.Range("A1").Offset(pp_i, var_i).Value
        Next pp_i
    Next var_i
    
    'helpful message boxes before moving on
    UserForm1.HelpText.Caption = "Column names = " & Join(vars_to_find, vbCrLf)
    Application.Wait (Now + TimeValue("0:00:01"))
    UserForm1.HelpText.Caption = "Person ids  = " & Join(pp_ids, vbCrLf)
    Application.Wait (Now + TimeValue("0:00:01"))
        
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 1B: READ IN VARIABLES FROM IMAGE DATA
    ' ---------------------------------------------------------------------------------------------------------
    Set ws = wb.Worksheets("image_data")
        
    ' different people might have different number of files
    Dim image_i, image_var_i, n_images_total As Integer
    Dim image_db() As Variant
    n_images_total = ws.Range("A2").End(xlDown).Row - 1
    ReDim image_db(n_images_total, 7) 'pp id, slide#, path, top, left, height, width
     
    UserForm1.HelpText.Caption = "Total # of images = " & n_images_total
    Application.Wait (Now + TimeValue("0:00:01"))
    
    For image_i = 1 To n_images_total
    
        DoEvents
        If stopit = True Then
            UserForm1.HelpText.Caption = "execution stopped. press 'x' to close"
            Exit Sub
        End If
    
        For image_var_i = 0 To 6
            image_db(image_i, image_var_i + 1) = ws.Range("A1").Offset(image_i, image_var_i).Value
        Next image_var_i
    Next image_i
    
    'Need to do validation:
    ' -- what if the slides are out of order
    ' -- or if person is non-consecutive
            
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 1C: READ IN VARIABLES FOR FORMATTING
    ' ---------------------------------------------------------------------------------------------------------
    Set ws = wb.Worksheets("formatting")
            
    Dim fmt_i, fmt_var_i, n_fmt_total, n_fmt_unique, unique_i As Integer
    Dim fmt_db(), fmt_unique() As Variant
    n_fmt_total = ws.Range("A2").End(xlDown).Row - 1
    ReDim fmt_unique(n_fmt_total, 3) 'fmt, start, stop
    ReDim fmt_db(n_fmt_total, 4) 'text, type, start ,len

    UserForm1.HelpText.Caption = "Total # of formatting = " & n_fmt_total
    Application.Wait (Now + TimeValue("0:00:01"))
    n_fmt_unique = 0
    unique_i = 1
    
    For fmt_i = 1 To n_fmt_total
    
        For fmt_var_i = 0 To 3
            fmt_db(fmt_i, fmt_var_i + 1) = ws.Range("A1").Offset(fmt_i, fmt_var_i).Value
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
            
            
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 2: REPLACE IN PPT AND SAVE NEW
    ' ---------------------------------------------------------------------------------------------------------
    Dim PowerPointApp As PowerPoint.Application
    Dim myPresentation As PowerPoint.Presentation
    
    'opens a powerpoint as read-only
    'its not terrible that this the behavior because it forces you to save as so you don't over-write the template
    UserForm1.HelpText.Caption = "Open ppt as read only"
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    'Dim outside the loop
    Dim sld As Slide
    Dim shp As PowerPoint.Shape
    Dim left_, right_ As String
    Dim search_text As String
    Dim textLoc, leftloc, rightloc As PowerPoint.TextRange
    Dim tablecell As PowerPoint.TextFrame
    Dim replace_val As Variant
    Dim nSlides, slide_i, table_i, table_j As Integer
    Dim sfolder As String
    Dim fullName As String
    Dim sFile As String
    Dim oPic As PowerPoint.Shape
    Dim image_path As String
    Dim loop_exit As Boolean
    Dim ppi, dpi As Double
    Dim image_top, image_left, image_height, image_width As Double
    Dim pos, start_fmt, len_fmt, pos_mod, fmt_unique_i As Integer
    Dim type_fmt As String
    
    image_i = 1
    left_ = Range("left_char").Value
    right_ = Range("right_char").Value

    ' Loop through for each person
    For pp_i = 1 To n_pp
        'doesn't show the the window
        'this continues to be this way as long as you don't say "activePresentation" or anything like that
        'this may need to change when we do images but we'll cross that bridge when we get there
        'this is the slowest part, is there anything faster we can do here?
        
        Set myPresentation = PowerPointApp.Presentations.Open(Range("ppt_template").Value, WithWindow:=msoFalse)
        'Set myPresentation = PowerPointApp.Presentations.Open(Range("ppt_template").Value)
        
        UserForm1.HelpText.Caption = "Find and replace"
        
        nSlides = myPresentation.Slides.Count

        sfolder = Range("dest_folder").Value & "\"
        sFile = pp_ids(pp_i)
        fullName = sfolder & sFile & " " & Format(Now, "yyyy-mm-dd hh-mm-ss")

        For slide_i = 1 To nSlides
        
            UserForm1.HelpText.Caption = "Person: " & pp_ids(pp_i) & ". Find and replace on Slide: " & slide_i
            Set sld = myPresentation.Slides(slide_i)
            
            
            ' updated 3/15/20 to just replace the specific range of text so formatting isn't messed up
            For Each shp In sld.Shapes
            '
                ' For replacing text data
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        
                        DoEvents
                        If stopit = True Then
                            UserForm1.HelpText.Caption = "execution stopped. press 'x' to close"
                            Exit Sub
                        End If
                        
                        ' For replacing text data, from key codes
                        For var_i = 1 To n_vars
                        
                            search_text = left_ & vars_to_find(var_i) & right_
                            
                            UserForm1.HelpText.Caption = "Person: " & pp_ids(pp_i) & _
                                    ". Find and replace on Slide: " & slide_i & _
                                    "; search text = " & search_text
                            
                            Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            While Not (textLoc Is Nothing)
                                replace_val = var_data(pp_i, var_i)
                                textLoc.Text = replace_val
                                Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            Wend
                            
                            ' give the function, the list of variables and the shape
                            ' doesnt return the shape, returns nothing,
                            ' as long as this doesnt pass by copy
                            ' so since its object oriented, we can assume this just passes a pointer
                            ' can i pull out a object with all the data
                            
                            
                        Next var_i
                        
                        ' for replacing formatting text
                        For fmt_unique_i = 1 To n_fmt_unique
                        
                            search_text = left_ & fmt_unique(fmt_unique_i, 1) & right_
                            Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            
                            UserForm1.HelpText.Caption = "Person: " & pp_ids(pp_i) & _
                                    ". Find and replace on Slide: " & slide_i & _
                                    "; format text = " & search_text
                            
                            
                            pos_mod = 1
                            
                            While Not (textLoc Is Nothing)

                                pos = InStr(pos_mod, shp.TextFrame.TextRange, search_text)
                                pos_mod = pos + 1
                                
                                ' so if you've found this text string still, that means it hasn't been touched at all
                                ' so go through each of the fmt steps for this and fix it, then unbracket and move on
                                
                                For fmt_i = fmt_unique(fmt_unique_i, 2) To fmt_unique(fmt_unique_i, 3)
                                    
                                    start_fmt = pos + fmt_db(fmt_i, 3)
                                    type_fmt = fmt_db(fmt_i, 2)
                                    len_fmt = fmt_db(fmt_i, 4)
                                    
                                    ' superscript
                                    If type_fmt = "superscript" Then
                                         shp.TextFrame.TextRange.Characters(start_fmt, len_fmt).Font.Superscript = True
                                    End If
                                    
                                    ' subscript
                                    If type_fmt = "subscript" Then
                                         shp.TextFrame.TextRange.Characters(start_fmt, len_fmt).Font.Subscript = True
                                    End If
                                    
                                    ' <<<<<<< FUTURE FIX
                                    ' this makes it so we don't have to say type_fmt == "superscript"
                                    'CallByName(shp.TextFrame.TextRange, "Characters(start_fmt, len_fmt).Font.Superscript", VbGet) = True
                                                               
                                Next fmt_i
                                
                                ' So at this point, this specific instance is totally fixed
                                Set leftloc = shp.TextFrame.TextRange.Find(left_, pos - 1)
                                leftloc.Text = ""
                                Set rightloc = shp.TextFrame.TextRange.Find(right_, pos + len_fmt - 1)
                                rightloc.Text = ""
                                                              
                                Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            
                            Wend
                            
                        Next fmt_unique_i
                        
                    End If
                End If
                
             '
                'replacing data in tables
                ' <<< need to make sure ther are no formatting characters in tables
                If shp.HasTable Then
                    
                    DoEvents
                    If stopit = True Then
                        UserForm1.HelpText.Caption = "execution stopped. press 'x' to close"
                        Exit Sub
                    End If
                    
                    For table_i = 1 To shp.Table.Rows.Count
                        For table_j = 1 To shp.Table.Columns.Count
                            Set tablecell = shp.Table.Rows.Item(table_i).Cells(table_j).Shape.TextFrame
                            
                            If tablecell.HasText Then
                                For var_i = 1 To n_vars
                                    search_text = left_ & vars_to_find(var_i) & right_
                                    Set textLoc = tablecell.TextRange.Find(search_text)
                                    While Not (textLoc Is Nothing)
                                        replace_val = var_data(pp_i, var_i)
                                        textLoc.Text = replace_val
                                        Set textLoc = tablecell.TextRange.Find(search_text)
                                    Wend
                                Next var_i
                            End If
                        Next table_j
                    Next table_i
                End If
              '
            Next shp
            
            'For inserting images
            'loop through the image db for each person
            loop_exit = False
            Do While image_i <= n_images_total And loop_exit = False
                
                If image_db(image_i, 1) = sFile And image_db(image_i, 2) = slide_i Then
                    
                    DoEvents
                    If stopit = True Then
                        UserForm1.HelpText.Caption = "execution stopped. press 'x' to close"
                        Exit Sub
                    End If
                                    
                    ppi = 72
                    dpi = 96
                                    
                    image_path = image_db(image_i, 3)
                    image_top = image_db(image_i, 4) * ppi
                    image_left = image_db(image_i, 5) * ppi
                    image_height = image_db(image_i, 6) * ppi
                    image_width = image_db(image_i, 7) * ppi
                                        
                    ' Check that image exists
                    If Dir(image_path) = "" Then
                        logger(logger_i) = "IMAGE FOR PERSON " & pp_i & " DOES NOT EXIST: " & image_path
                        logger_i = logger_i + 1
                    Else
                        Set oPic = sld.Shapes.AddPicture(image_path, False, True, image_left, image_top, image_width, image_height)
                    End If
                    
                    image_i = image_i + 1
                Else
                    loop_exit = True
                End If
                        
            Loop

            
            
        Next slide_i
        
        'export to pdf or ppt
        UserForm1.HelpText.Caption = "Save as " & fullName & "." & Range("output_as").Value
        If Range("output_as").Value = "pdf" Then
            myPresentation.ExportAsFixedFormat fullName & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentScreen
        End If
        If Range("output_as").Value = "ppt" Then
            myPresentation.SaveAs Filename:=fullName & ".pptx"
        End If
        
        progress pp_i / n_pp * 100
        
    Next pp_i
    
    ' Logger export and messaging
    If logger_i > 1 Then
        
        For logprint_i = 1 To (logger_i - 1)
            
            Sheets("ErrorLog").Cells(1 + logprint_i, 2).Value = logger(logprint_i)
            
        Next logprint_i
            
        MsgBox ("CHECK ERROR LOG")
        
    End If
            
    ' PowerPointApp.Quit
    Application.ScreenUpdating = True
    UserForm1.HelpText.Caption = "Finished"
    
End Sub

