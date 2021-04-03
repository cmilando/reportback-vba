Attribute VB_Name = "main"
'@Folder "VBAProject"
Option Explicit
Public stopit As Boolean

' ! ===========================================================================
Sub main()

    Progress_bar.Show
    
End Sub

' ! ===========================================================================
Sub progress(pctCompl As Single)

    Progress_bar.Text.Caption = pctCompl & "% Completed"
    Progress_bar.Bar.Width = pctCompl * 2
    
    DoEvents
    
    If stopit = True Then
        Progress_bar.HelpText.Caption = "execution stopped. press 'x' to close"
        Exit Sub
    End If

End Sub

' ! ===========================================================================
Sub test_class()

    Dim p_data As New Participant_data
    
    Debug.Print p_data.n
    
    Debug.Print "done"

End Sub

' ! ===========================================================================
Sub find_and_replace()
       
    'need to add the powerpoint library to tools > references
    stopit = False
    Application.ScreenUpdating = False
    Progress_bar.HelpText.Caption = "Starting"
    
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 0: CLEAR ERROR LOG, VALIDATION CHECKS
    ' ---------------------------------------------------------------------------------------------------------
        
    ' Clear the error log page, and set up the logger
    Dim Logger() As Variant                     ' re-dimable logger
    Dim logger_i, logprint_i As Integer         ' iterate through logger
    Const logger_size As Integer = 1000 ' Logger size
    ReDim Logger(logger_size)
    logger_i = 1
    logprint_i = 1
    Sheets("ErrorLog").Cells.Clear

    ' ---------------------------------------------------------------------------------------------------------
    ' PART 1: READ IN VARIABLES FROM TEXT DATA
    ' ---------------------------------------------------------------------------------------------------------
    
    Dim p_data As New Participant_data
    
    p_data.fill_vars
    p_data.fill_image_data
    p_data.fill_formatting_chars
            
    ' ---------------------------------------------------------------------------------------------------------
    ' PART 2: REPLACE IN PPT AND SAVE NEW
    ' ---------------------------------------------------------------------------------------------------------
    Dim PowerPointApp As PowerPoint.Application
    Dim myPresentation As PowerPoint.Presentation
    
    'opens a powerpoint as read-only
    'its not terrible that this the behavior because it forces you to save as so you don't over-write the template
    Progress_bar.HelpText.Caption = "Open ppt as read only"
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
    Dim pos, start_fmt, len_fmt, pos_mod As Integer
    Dim type_fmt As String
    
    Dim image_i, pp_i, key_i, fmt_unique_i, fmt_i As Integer
    
    image_i = 1
    left_ = Range("left_char").value
    right_ = Range("right_char").value
    
    ' This loop covers
    ' - textframes > keycodes
    ' - textframes > formatting text
    ' - text in tables
    ' - images
    
    ' Loop through for each person
    For pp_i = 1 To p_data.n_persons
        'doesn't show the the window
        'this continues to be this way as long as you don't say "activePresentation" or anything like that
        'this may need to change when we do images but we'll cross that bridge when we get there
        'this is the slowest part, is there anything faster we can do here?
        
        Set myPresentation = PowerPointApp.Presentations.Open(Range("ppt_template").value, WithWindow:=msoFalse)
        'Set myPresentation = PowerPointApp.Presentations.Open(Range("ppt_template").Value)
        
        Progress_bar.HelpText.Caption = "Find and replace"
        
        nSlides = myPresentation.Slides.Count

        sfolder = Range("dest_folder").value & "\"
        sFile = p_data.person_ids(pp_i)
        
        For slide_i = 1 To nSlides
        
            Progress_bar.HelpText.Caption = "Person: " & sFile & ". Find and replace on Slide: " & slide_i
            Set sld = myPresentation.Slides(slide_i)
            
            
            ' updated 3/15/20 to just replace the specific range of text so formatting isn't messed up
            For Each shp In sld.Shapes
            '
                ' For replacing text data
                If shp.HasTextFrame Then
                
                    If shp.TextFrame.HasText Then
                        
                        DoEvents
                        If stopit = True Then
                            MsgBox ("Execution stopped")
                            End
                        End If
                        
                        ' For replacing text data, from key codes
                        For key_i = 1 To p_data.n_keys
                        
                            search_text = left_ & p_data.keys(key_i) & right_
                            
                            Progress_bar.HelpText.Caption = "Person: " & sFile & _
                                    ". Find and replace on Slide: " & slide_i & _
                                    "; search text = " & search_text
                            
                            Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            
                            While Not (textLoc Is Nothing)
                                replace_val = p_data.var_data(pp_i, key_i)
                                textLoc.Text = replace_val
                                Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            Wend
                            
                            ' give the function, the list of variables and the shape
                            ' doesnt return the shape, returns nothing,
                            ' as long as this doesnt pass by copy
                            ' so since its object oriented, we can assume this just passes a pointer
                            ' can i pull out a object with all the data
                            
                            
                        Next key_i
                        
                        ' for replacing formatting text
                        For fmt_unique_i = 1 To p_data.n_fmt
                        
                            search_text = left_ & p_data.fmt_u(fmt_unique_i, 1) & right_
                            Set textLoc = shp.TextFrame.TextRange.Find(search_text)
                            
                            Progress_bar.HelpText.Caption = "Person: " & sFile & _
                                    ". Find and replace on Slide: " & slide_i & _
                                    "; format text = " & search_text
                            
                            
                            pos_mod = 1
                            
                            While Not (textLoc Is Nothing)

                                pos = InStr(pos_mod, shp.TextFrame.TextRange, search_text)
                                pos_mod = pos + 1
                                
                                ' so if you've found this text string still, that means it hasn't been touched at all
                                ' so go through each of the fmt steps for this and fix it, then unbracket and move on
                                
                                For fmt_i = p_data.fmt_u(fmt_unique_i, 2) To p_data.fmt_u(fmt_unique_i, 3)
                                    
                                    start_fmt = pos + p_data.fmt_data(fmt_i, 3)
                                    type_fmt = p_data.fmt_data(fmt_i, 2)
                                    len_fmt = p_data.fmt_data(fmt_i, 4)
                                    
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
                        MsgBox ("Execution stopped")
                        End
                    End If
                    
                    For table_i = 1 To shp.Table.Rows.Count
                    
                        For table_j = 1 To shp.Table.Columns.Count
                        
                            Set tablecell = shp.Table.Rows.Item(table_i).Cells(table_j).Shape.TextFrame
                            
                            If tablecell.HasText Then
                                For key_i = 1 To p_data.n_keys
                                
                                    search_text = left_ & p_data.keys(key_i) & right_
                                    Set textLoc = tablecell.TextRange.Find(search_text)
                                    
                                    While Not (textLoc Is Nothing)
                                        replace_val = p_data.var_data(pp_i, key_i)
                                        textLoc.Text = replace_val
                                        Set textLoc = tablecell.TextRange.Find(search_text)
                                    Wend
                                    
                                Next key_i
                            End If
                        Next table_j
                    Next table_i
                End If
              '
            Next shp
            
            'For inserting images
            ' loop through the image db for each person
            ' probably better to do this as a collection ...
            ' or just check for it in the validation steps
            loop_exit = False
            Do While image_i <= p_data.n_images And loop_exit = False
                
                If p_data.images(image_i, 1) = sFile And p_data.images(image_i, 2) = slide_i Then
                    
                    DoEvents
                    If stopit = True Then
                        MsgBox ("Execution stopped")
                        End
                    End If
                                    
                    ppi = 72
                    dpi = 96
                                    
                    image_path = p_data.images(image_i, 3)
                    image_top = p_data.images(image_i, 4) * ppi
                    image_left = p_data.images(image_i, 5) * ppi
                    image_height = p_data.images(image_i, 6) * ppi
                    image_width = p_data.images(image_i, 7) * ppi
                                        
                    ' Check that image exists
                    If Dir(image_path) = "" Then
                        Logger(logger_i) = "IMAGE FOR PERSON " & pp_i & " DOES NOT EXIST: " & image_path
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
        If Range("output_suffix").value = "date" Then
            fullName = sfolder & sFile & " " & Format(Now, "yyyy-mm-dd hh-mm-ss")
        End If
        If Range("output_suffix").value = "none" Then
            fullName = sfolder & sFile
        End If
        
        Progress_bar.HelpText.Caption = "Save as " & fullName & "." & Range("output_as").value
        If Range("output_as").value = "pdf" Then
            myPresentation.ExportAsFixedFormat fullName & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentScreen
        End If
        If Range("output_as").value = "ppt" Then
            myPresentation.SaveAs Filename:=fullName & ".pptx"
        End If
        
        ' update progress bar
        progress pp_i / p_data.n_persons * 100
        
    Next pp_i
    
    ' Logger export and messaging
    If logger_i > 1 Then
        
        For logprint_i = 1 To (logger_i - 1)
            
            Sheets("ErrorLog").Cells(1 + logprint_i, 2).value = Logger(logprint_i)
            
        Next logprint_i
            
        MsgBox ("CHECK ERROR LOG")
        
    End If
            
    ' PowerPointApp.Quit
    Application.ScreenUpdating = True
    Progress_bar.HelpText.Caption = "Finished"
    Progress_bar.CommandButton1.Visible = False
    Progress_bar.Label1.Visible = False
    
End Sub


