Attribute VB_Name = "main"
'@Folder "VBAProject"
Option Explicit
Public stopit As Boolean

' ! ===========================================================================
Sub find_and_replace()
       
    'need to add the powerpoint library to tools > references
    stopit = False
    Application.ScreenUpdating = False
    Progress_bar.HelpText.Caption = "Starting"
       
    ' -------------------------------------------------------------------------
    ' PART 1: READ IN PARTICIPANT DATA
    ' -------------------------------------------------------------------------
    Progress_bar.HelpText.Caption = "Open excel data workbook"
    
    Dim p_data As New Participant_data
        
    Call p_data.fill_vars
    
    If Range("use_image_data").value = "yes" Then
        Call p_data.fill_image_data
    End If
    
    If Range("use_formatting_data").value = "yes" Then
        Call p_data.fill_formatting_chars
    End If
    
    ' -------------------------------------------------------------------------
    ' PART 2: REPLACE IN PPT AND SAVE NEW
    ' -------------------------------------------------------------------------
    Dim PowerPointApp As PowerPoint.Application
    Dim myPresentation As PowerPoint.Presentation
    
    ' opens a powerpoint as read-only
    Progress_bar.HelpText.Caption = "Open ppt as read only"
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    'Dim outside the loop
    Dim sld As slide
    Dim shp As PowerPoint.Shape
    Dim left_ As String
    Dim right_ As String
    Dim dest_folder As String
    Dim output_suffix As String
    Dim output_type As String
    Dim template_file As String
    Dim fullName As String
    Dim p_id As String
    
    ' iterators
    Dim nSlides As Integer
    Dim slide_i As Integer
    Dim pp_i As Integer

    left_ = Range("left_char").value
    right_ = Range("right_char").value
    dest_folder = Range("dest_folder").value & "\"
    output_suffix = Range("output_suffix").value
    output_type = Range("output_as").value
    template_file = Range("template").value
       
    ' -------------------------------------------------------------------------
    ' Loop through for each person
    For pp_i = 1 To p_data.n_persons
        
        ' open a new version of the presentation
        Set myPresentation = _
               PowerPointApp.Presentations.Open(template_file, _
               WithWindow:=msoFalse)
        
        Progress_bar.HelpText.Caption = "Find and replace"
        
        nSlides = myPresentation.Slides.Count

        p_id = p_data.person_ids(pp_i)

        ' loop through each slide
        For slide_i = 1 To nSlides
        
            Progress_bar.HelpText.Caption = "Person: " & p_id & _
                          ". Find and replace on Slide: " & slide_i
                          
            Set sld = myPresentation.Slides(slide_i)
            
            ' insert text into text boxes and tables
            For Each shp In sld.Shapes
                
                Call replace_text_in_shape(shp, p_data, slide_i, pp_i, _
                                            left_, right_)
                
                Call replace_text_in_table(shp, p_data, slide_i, pp_i, _
                                            left_, right_)
                
            Next shp
            
            ' insert images onto slides
            If Range("use_image_data").value = "yes" Then
                Call insert_images(sld, p_data, slide_i, pp_i)
            End If

        Next slide_i
        
        'export to pdf or ppt
        Call export_to_file(p_data, myPresentation, output_suffix, _
                            output_type, dest_folder, pp_i)
                
        ' update progress bar
        progress pp_i / p_data.n_persons * 100
        
    Next pp_i
    ' End master loop
    
    ' -------------------------------------------------------------------------
    ' Logger export and messaging
    Call p_data.print_logger
            
    ' Close up
    Application.ScreenUpdating = True
    Progress_bar.HelpText.Caption = "Finished"
    Progress_bar.CommandButton1.Visible = False
    Progress_bar.Label1.Visible = False
    Application.Wait (Now + TimeValue("0:00:01"))
    Progress_bar.Hide
    
End Sub

' ! ===========================================================================
Sub main()
    
    progress 0
    Progress_bar.Show
    
End Sub

' ! ===========================================================================
Sub progress(pctCompl As Single)

    Progress_bar.Text.Caption = pctCompl & "% Completed"
    Progress_bar.Bar.width = pctCompl * 2
        
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
Sub replace_text_placeholders(obj As Object, _
                                 p_data As Participant_data, _
                                 slide_i As Integer, _
                                 pp_i As Integer, _
                                 left_ As String, _
                                 right_ As String)
    Dim key_i As Integer
    Dim search_text As String
    Dim textLoc  As PowerPoint.TextRange
    Dim p_id As String
    Dim replace_val As Variant
    
    p_id = p_data.person_ids(pp_i)
    
    For key_i = 1 To p_data.n_keys
    
        search_text = left_ & p_data.keys(key_i) & right_
        
        Progress_bar.HelpText.Caption = "Person: " & p_id & _
                ". Find and replace on Slide: " & slide_i & _
                "; search text = " & search_text
        
        Set textLoc = obj.TextRange.Find(search_text)
        
        While Not (textLoc Is Nothing)
            replace_val = p_data.var_data(pp_i, key_i)
            textLoc.Text = replace_val
            Set textLoc = obj.TextRange.Find(search_text)
        Wend
        
    Next key_i

End Sub

' ! ===========================================================================
Sub replace_formatting_placeholders(obj As Object, _
                                 p_data As Participant_data, _
                                 slide_i As Integer, _
                                 pp_i As Integer, _
                                 left_ As String, _
                                 right_ As String)
    
    Dim key_i As Integer
    Dim search_text As String
    Dim textLoc As PowerPoint.TextRange
    Dim rightLoc As PowerPoint.TextRange
    Dim leftLoc As PowerPoint.TextRange
    Dim p_id As String
    Dim replace_val As Variant
    Dim pos As Integer
    Dim start_fmt As Integer
    Dim len_fmt As Integer
    Dim pos_mod As Integer
    Dim type_fmt As String
    Dim fmt_unique_i As Integer
    Dim fmt_i As Integer
    
    For fmt_unique_i = 1 To p_data.n_fmt
                            
        search_text = left_ & p_data.fmt_u(fmt_unique_i, 1) & right_
        Set textLoc = obj.TextRange.Find(search_text)
        
        Progress_bar.HelpText.Caption = "Person: " & p_id & _
                ". Find and replace on Slide: " & slide_i & _
                "; format text = " & search_text
        
        pos_mod = 1
        
        While Not (textLoc Is Nothing)
    
            pos = InStr(pos_mod, obj.TextRange, search_text)
            pos_mod = pos + 1
            
            ' so if you've found this text string still, that means it hasn't been touched at all
            ' so go through each of the fmt steps for this and fix it, then unbracket and move on
            
            For fmt_i = p_data.fmt_u(fmt_unique_i, 2) To p_data.fmt_u(fmt_unique_i, 3)
                
                start_fmt = pos + p_data.fmt_data(fmt_i, 3)
                type_fmt = p_data.fmt_data(fmt_i, 2)
                len_fmt = p_data.fmt_data(fmt_i, 4)
                
                ' superscript
                If type_fmt = "superscript" Then
                     obj.TextRange.Characters(start_fmt, len_fmt).Font.Superscript = True
                End If
                
                ' subscript
                If type_fmt = "subscript" Then
                     obj.TextRange.Characters(start_fmt, len_fmt).Font.Subscript = True
                End If
                
                ' <<<<<<< FUTURE FIX
                ' this makes it so we don't have to say type_fmt == "superscript"
                'CallByName(obj.TextRange, "Characters(start_fmt, len_fmt).Font.Superscript", VbGet) = True
                                           
            Next fmt_i
            
            ' So at this point, this specific instance is totally fixed
            Set leftLoc = obj.TextRange.Find(left_, pos - 1)
            leftLoc.Text = ""
            Set rightLoc = obj.TextRange.Find(right_, pos + len_fmt - 1)
            rightLoc.Text = ""
                                          
            Set textLoc = obj.TextRange.Find(search_text)
        
        Wend
        
    Next fmt_unique_i


End Sub

' ! ===========================================================================
Sub replace_text_in_shape(shp As PowerPoint.Shape, p_data As Participant_data, _
                                 slide_i As Integer, _
                                 pp_i As Integer, _
                                 left_ As String, _
                                 right_ As String)

    ' For replacing text data
    If shp.HasTextFrame Then
        
        If shp.TextFrame.HasText Then
            
            Call check_cancel
            
            ' For replacing text data, from key codes
            Call replace_text_placeholders(shp.TextFrame, _
                p_data, slide_i, pp_i, left_, right_)
            
            ' for replacing formatting text
            If Range("use_formatting_data").value = "yes" Then
                Call replace_formatting_placeholders(shp.TextFrame, _
                    p_data, slide_i, pp_i, left_, right_)
            End If
                
        End If
        
    End If

End Sub

' ! ===========================================================================
Sub replace_text_in_table(shp As PowerPoint.Shape, p_data As Participant_data, _
                                 slide_i As Integer, _
                                 pp_i As Integer, _
                                 left_ As String, _
                                 right_ As String)

    ' table vars
    Dim tablecell As PowerPoint.TextFrame
    Dim table_i As Integer
    Dim table_j As Integer
    
    
    'replacing data in tables
    If shp.HasTable Then
                            
        Call check_cancel
        
        For table_i = 1 To shp.Table.Rows.Count
        
            For table_j = 1 To shp.Table.Columns.Count
            
                Set tablecell = shp.Table.Rows.Item(table_i).Cells(table_j).Shape.TextFrame
                
                If tablecell.HasText Then
                
                    ' For replacing text data, from key codes
                    Call replace_text_placeholders(tablecell, _
                        p_data, slide_i, pp_i, left_, right_)
            
                    ' for replacing formatting text
                    If Range("use_formatting_data").value = "yes" Then
                        Call replace_formatting_placeholders(tablecell, _
                            p_data, slide_i, pp_i, left_, right_)
                    End If
                
                End If
                
            Next table_j
            
        Next table_i
        
    End If

End Sub

' ! ===========================================================================
Sub insert_images(sld As PowerPoint.slide, p_data As Participant_data, _
                                 slide_i As Integer, pp_i As Integer)

    ' loop through the image db for each person
    ' probably better to do this as a collection ...
    ' or just check for it in the validation steps
    
    Dim oPic As PowerPoint.Shape
    Dim image_path As String
    Dim loop_exit As Boolean
    Dim ppi As Double
    Dim dpi As Double
    Dim image_top  As Double
    Dim image_left  As Double
    Dim image_height As Double
    Dim image_width As Double
    Dim p_id As String
    Dim key As String
    Dim image_i As Variant
        
    loop_exit = False
    
    p_id = p_data.person_ids(pp_i)
    key = p_id & "_" & slide_i
    
    If Exists(p_data.images, key) Then
    
        ' For each image in key add it to this slide
        For Each image_i In p_data.images(key)
            
            ppi = 72
            dpi = 96
                            
            image_path = image_i.path
            image_top = image_i.top * ppi
            image_left = image_i.left * ppi
            image_height = image_i.height * ppi
            image_width = image_i.width * ppi
                                
            ' This assumes all images exist!!
            Set oPic = sld.Shapes.AddPicture(image_path, False, _
                       True, image_left, image_top, image_width, image_height)
            
        Next image_i
    
    End If

End Sub

' ! ===========================================================================
Sub export_to_file(p_data As Participant_data, _
                   myPresentation As PowerPoint.Presentation, _
                   output_suffix As String, _
                   output_type As String, _
                   dest_folder As String, _
                   pp_i As Integer)
    
    Dim fullName As String
    Dim p_id As String
    
    p_id = p_data.person_ids(pp_i)
    
    ' add a date suffix
    If output_suffix = "date" Then
        fullName = dest_folder & p_id & " " & _
                   Format(Now, "yyyy-mm-dd hh-mm-ss")
    End If
    
    ' don't add a suffix and overwrite
    If output_suffix = "none" Then
        fullName = dest_folder & p_id
    End If
    
    Progress_bar.HelpText.Caption = "Save as " & fullName & _
                                    "." & output_type
    
    ' export as pdf
    If output_type = "pdf" Then
        myPresentation.ExportAsFixedFormat fullName & ".pdf", _
                  ppFixedFormatTypePDF, ppFixedFormatIntentScreen
    End If
    
    ' export as powerpoint
    If output_type = "ppt" Then
        myPresentation.SaveAs Filename:=fullName & ".pptx"
    End If

End Sub

' ! ===========================================================================
Function Exists(coll As Collection, key As String) As Boolean

    On Error GoTo EH

    IsObject (coll.Item(key))
    
    Exists = True
EH:
End Function
