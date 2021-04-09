VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Image_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Public pp As String
Public slide As Integer
Public path As String
Public top  As Double
Public left  As Double
Public height As Double
Public width As Double

Sub set_values(pp_i As Variant, slide_i As Variant, _
            image_path As Variant, image_top As Variant, _
            image_left As Variant, image_height As Variant, _
            image_width As Variant)
    
    
        ' Slide # has to be numeric
        If Not value_in_set(VarType(slide_i), Array(2, 3, 4, 5)) Then
            MsgBox "ERROR: slide # should be numeric - see row " & image_i & _
                   " with value " & slide_i
            End
        End If
               
        ' Valid filepath
        If Not IsFile(image_path) Then
            MsgBox "ERROR: Image on line " & image_i & _
                    " does not exist! Path: " & image_path
            End
        End If
        
        ' top
        If Not value_in_set(VarType(top), Array(2, 3, 4, 5)) Then
            MsgBox "ERROR: 'top' should be numeric - see row " & image_i & _
                   " with value " & top
            End
        End If

        ' left
        If Not value_in_set(VarType(left), Array(2, 3, 4, 5)) Then
            MsgBox "ERROR: 'left' should be numeric - see row " & image_i & _
                   " with value " & left
            End
        End If

        ' height
        If Not value_in_set(VarType(height), Array(2, 3, 4, 5)) Then
            MsgBox "ERROR: 'height' should be numeric - see row " & image_i & _
                   " with value " & height
            End
        End If

        ' width
        If Not value_in_set(VarType(width), Array(2, 3, 4, 5)) Then
            MsgBox "ERROR: 'width' should be numeric - see row " & image_i & _
                   " with value " & width
            End
        End If
    
    pp = pp_i
    slide = CInt(slide_i)
    path = image_path
    top = image_top
    left = image_left
    height = image_height
    width = image_width

End Sub
' ! ===========================================================================
Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function
' ! ===========================================================================
Function value_in_set(strIn As Variant, arrList As Variant)
    
    value_in_set = Not IsError(Application.Match(strIn, arrList, 0))
        
End Function
