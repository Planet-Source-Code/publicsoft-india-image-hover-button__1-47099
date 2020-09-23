VERSION 5.00
Begin VB.UserControl ImageButton 
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ScaleHeight     =   191
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ToolboxBitmap   =   "imgButton.ctx":0000
   Begin VB.Image img 
      Height          =   1815
      Left            =   120
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "ImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOVER IMAGE BUTTON ACTIVEX
'By Cyril M Gupta
'cyril@cyrilgupta.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'API CALLS
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const LF_FACESIZE = 32

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORD_ELLIPSIS = &H40000



Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_3DDKSHADOW = 21
Private Const COLOR_3DLIGHT = 22
Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24

Private Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


Private Const PS_SOLID = 0
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_DOT = 2                     '  .......
Private Const PS_DASH = 1                    '  -------
Private Const PS_ENDCAP_FLAT = &H200

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long


Enum PicStates
    picNothing = 0
    picDown = 1
    picHover = 2
    picNorm = 3
End Enum

Dim PicState As PicStates

Dim m_NormalPic As New StdPicture 'Normal picture
Dim m_HoverPic As New StdPicture 'Hover picture
Dim m_DownPic As New StdPicture 'Down picture
Dim m_MouseInside As Boolean

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseOut()
Public Event Resize()

Private Sub img_Click()
RaiseEvent Click
End Sub

Private Sub img_DblClick()
RaiseEvent DblClick
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOW IT WORKS
'Everything else in this control should be self
'explanatory. The only thing that matters is the
'mousemove function of the image and the user control.
'When the control first detects the mousemove, it calls
'the Getcapture function to find if the mouse is captured
'by any window, if it is not, it calls the setcapture
'function to capture the mouse by the user control window
'this way no matter where the mouse curser is, our user
'control will recieve its mesage, when the user control
'detects that the mouse is outside its borders it releases
'capture.
'
'Took me 3 hours to make, with some sample code to work on
'Cheerios
'Cyril M Gupta
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub img_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)

If Button = vbLeftButton Then
    Exit Sub
End If

If GetCapture() <> UserControl.hwnd Then
    SetCapture (UserControl.hwnd)
    If Not img.Picture = HoverPic Then
        img.Picture = HoverPic
    End If
Else
    Dim pt As POINTAPI
    pt.x = x
    pt.y = y
    ClientToScreen UserControl.hwnd, pt
    If WindowFromPoint(pt.x, pt.y) <> UserControl.hwnd Then
        Refresh
        If Button <> vbLeftButton Then
            ReleaseCapture
            img.Picture = NormalPic
            RaiseEvent MouseOut
        End If
        End If
End If
End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
RaiseEvent Click
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
If Not NormalPic Is Nothing Then
    Set img.Picture = NormalPic
End If
img.Top = 0
img.Left = 0
End Sub

Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
img.Picture = DownPic
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub usercontrol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)

If Button = vbLeftButton Then
    Exit Sub
End If

If GetCapture() <> UserControl.hwnd Then
    SetCapture (UserControl.hwnd)
    If Not img.Picture = HoverPic Then
        img.Picture = HoverPic
        m_MouseInside = True
    End If
Else
    Dim pt As POINTAPI
    pt.x = x
    pt.y = y
    ClientToScreen UserControl.hwnd, pt
    If WindowFromPoint(pt.x, pt.y) <> UserControl.hwnd Then
        Refresh
        If Button <> vbLeftButton Then
            ReleaseCapture
            img.Picture = NormalPic
            m_MouseInside = False
            RaiseEvent MouseOut
        End If
        End If
End If
End Sub

Private Sub usercontrol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
img.Picture = HoverPic
RaiseEvent MouseUp(Button, Shift, x, y)
RaiseEvent Click
End Sub

Public Property Get NormalPic() As StdPicture
Set NormalPic = m_NormalPic
End Property

Public Property Set NormalPic(vNewPic As StdPicture)
Set m_NormalPic = vNewPic
PropertyChanged "NormalPic"
img.Picture = NormalPic
End Property

Public Property Get DownPic() As StdPicture
Set DownPic = m_DownPic
End Property

Public Property Set DownPic(vNewPic As StdPicture)
Set m_DownPic = vNewPic
PropertyChanged "DownPic"
End Property

Public Property Get HoverPic() As StdPicture
Set HoverPic = m_HoverPic
End Property

Public Property Set HoverPic(vNewPic As StdPicture)
Set m_HoverPic = vNewPic
PropertyChanged "HoverPic"
End Property

Private Sub UserControl_Paint()
img.Picture = NormalPic
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set m_NormalPic = PropBag.ReadProperty("NormalPic", Nothing)
Set m_DownPic = PropBag.ReadProperty("DownPic", Nothing)
Set m_HoverPic = PropBag.ReadProperty("HoverPic", Nothing)
img.Stretch = PropBag.ReadProperty("Stretch", False)
End Sub

Private Sub UserControl_Resize()
RaiseEvent Resize
img.Width = UserControl.ScaleWidth
img.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "NormalPic", m_NormalPic, 0
PropBag.WriteProperty "DownPic", m_DownPic, 0
PropBag.WriteProperty "HoverPic", m_HoverPic, 0
PropBag.WriteProperty "Stretch", img.Stretch
End Sub

Public Property Get Stretch() As Boolean
Stretch = img.Stretch
End Property

Public Property Let Stretch(vNewValue As Boolean)
img.Stretch = vNewValue
PropertyChanged "Stretch"
End Property

Public Property Get CurPicture() As PicStates
If img.Picture = 0 Then
    CurPicture = picNothing
ElseIf img.Picture = NormalPic Then
    CurPicture = picNorm
ElseIf img.Picture = DownPic Then
    CurPicture = picDown
ElseIf img.Picture = HoverPic Then
    CurPicture = picHover
End If
End Property

Public Property Get MouseInside() As Boolean
MouseInside = m_MouseInside
End Property
