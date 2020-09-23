VERSION 5.00
Begin VB.UserControl ucAutoResize 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucAutoResize.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "ucAutoResize.ctx":0C44
End
Attribute VB_Name = "ucAutoResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   ucAutoResize.ctl
'

'   Based on the idea of the very good submission from Hamed Oveisi (thx man!)
'   (look/vote at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=49740&lngWId=1)
'   I tried to improve/change some things ('Refactoring' is the trendy word for ;)  ):
'
'
'   *  More possibilities in resizing (look at the 3D lines (frames) ).
'   *  No more flickering when forms gets too small.
'   *  No more call in the Form_Resize event neccessary. (In fact you didn't need ANY code.)
'      (Ups! Just check out Hamed's last update. He did the same ;) )
'   *  The tag value still can be used for "standard" purposes. (But you will need a (very) litte
'         change in your code, sorry ;( )
'   *  Its faster.
'   *  Handling is easier/more straight forward. (Only two simple digits in tag value to enter, not four)
'   *  Prepared for handling 'Lines', too. (Not done by me - I don't use lines ;) )
'   *  A var naming convention is used, so code is easier to read/modify.
'   *  Demo and description extended.
'   *  ...
'
'   All  -(C) Copyleft on 11/10/2003 - Light Templer (LiTe)
'
'
'   Usage:      Simply put on every control you want to automaticly be resized by ucAutoResize
'               three additional digits to the end of the TAG value.
'
'               1 -  '|' as delimiter to the rest. (Can be changed - look for Const TAG_DELIMITER )
'               2 -  'L' or 'R' or '-'
'               3 -  'T' or 'B' or '-'
'
'               e.g.:       'Normal tag value1 |LB'
'                           'Val1, Val2, Val3|R-'
'                           '|-T'
'
'               'L' :       LEFT border of the control will be moved when form is resized.
'                           Width of control will not be changed.
'
'               'R' :       RIGHT border of the control will be moved when form is resized.
'                           Width of control will be changed.
'
'
'               'T' :       TOP border of the control will be moved when form is resized.
'                           Height of control will not be changed.
'
'               'B' :       BOTTOM border of the control will be moved when form is resized.
'                           Height of control will be changed.
'
'               '-' :       Don't change anything on this direction.
'
'   Thats all! Simple and easy to remember!
'   Have fun with it and kind regards
'
'   LiTe
'

Option Explicit

Private Const TAG_DELIMITER = "|"               ' Change to your needs! e.g.  "~"  or  ","  or any other.


Private lLastHeight     As Long
Private lLastWidth      As Long

Private lSizeMinX       As Long
Private lSizeMinY       As Long

Private WithEvents oFrm As Form
Attribute oFrm.VB_VarHelpID = -1
'
'
'


Private Sub oFrm_Resize()
    ' We can easily catch the form's Resize event here. Thats imho a nice feature of VB!
        
    Dim lHeightChange   As Long
    Dim lWidthChange    As Long
    Dim lControls       As Long
    Dim i               As Long
    Dim lPos            As Long
    Dim sTag            As String
    
    On Local Error Resume Next
               
    With oFrm
        If .WindowState = vbMinimized Then
            
            Exit Sub
        End If
        
        lHeightChange = .Height - lLastHeight
        lWidthChange = .Width - lLastWidth
        
        If lLastHeight And lLastWidth Then                  ' Not on first time
            lControls = .Controls.Count - 1
            For i = 0 To lControls
                If Not TypeOf .Controls(i) Is Line Then     ' (Lines have X1/Y1,X2/Y2 ... :((( )
                    
                    sTag = vbNullString                     ' On controls without a TAG property, Error
                    sTag = .Controls(i).Tag                 ' Resume Next skips this line.
                    
                    ' Get the part of the TAG value behind the delimiter
                    lPos = InStr(1, sTag, TAG_DELIMITER)
                    If lPos And lPos + 2 = Len(sTag) Then
                        sTag = Mid$(sTag, lPos + 1)
                        
                        If .Width >= lSizeMinX Then
                            If Left$(sTag, 1) = "L" Then
                                .Controls(i).Left = .Controls(i).Left + lWidthChange
                            ElseIf Left$(sTag, 1) = "R" Then
                                .Controls(i).Width = .Controls(i).Width + lWidthChange
                            End If
                        End If
                        
                        If .Height >= lSizeMinY Then
                            If Right$(sTag, 1) = "T" Then
                                .Controls(i).Top = .Controls(i).Top + lHeightChange
                            ElseIf Right$(sTag, 1) = "B" Or Right$(sTag, 1) = "A" Then
                                .Controls(i).Height = .Controls(i).Height + lHeightChange
                            End If
                        End If
                        
                    End If
                Else
                    ' Insert code for 'Line' control handling here (if you really need it ... ;) )
                End If
            Next i
        End If
        
        If .Width >= lSizeMinX Then
            lLastWidth = .Width
        End If
        If .Height >= lSizeMinY Then
            lLastHeight = .Height
        End If
    End With
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Local Error Resume Next
    
    ' We are in RunTime-Mode ?   (At design-time we don't have resize events :( , no matter what type
    '                             of form we use (Form, VBForm, ...))
    If UserControl.Ambient.UserMode = True Then
         ' So it's safe to get a ref to the parent of the control (must be a standard VB form!)
        Set oFrm = UserControl.Extender.Parent
        If Not oFrm Is Nothing Then
            lSizeMinX = oFrm.Width
            lSizeMinY = oFrm.Height
        End If
    End If

End Sub

Private Sub UserControl_Resize()
    
   ' Fixed to image's size
   UserControl.Width = 480
   UserControl.Height = 480
   
End Sub


' #*#

