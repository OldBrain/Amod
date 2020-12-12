VERSION 5.00
Begin VB.UserControl GradButton 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   167
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1980
      Tag             =   "Down"
      Top             =   0
   End
   Begin VB.PictureBox pctBut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GradButton"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "GradButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'BUTTON
'Created by 13GHOST
'mailto:13GHOST@mail.ru
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, bits As Any, BitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
'++++++++++++++++++++++++

Dim R1 As Integer, G1 As Integer, B1 As Integer

Private Type GRADIENTSTRUCT
    BaseColor As Long
    FinalColor As Long
    Percentage As Single
End Type

Private Type POINT_API
    x As Long
    y As Long
End Type

Private Type RGB
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Private Type RGBTHREE '
    rgbtBlue As Integer
    rgbtGreen As Integer
    rgbtRed As Integer
End Type
Private H, k
Private PixelRGB() As Long
Private lpGradient As GRADIENTSTRUCT

Private StartColor As Long ' ��������� ����
Private FinishColor As Long ' �������� ����

Private SizeX As Integer, SizeY As Integer
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Private hBrush As Long
Private lpBrush As LOGBRUSH

Private m_pBrush As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private lpRect As RECT
'������
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Function GetGradColor(lpGradient As GRADIENTSTRUCT) As Long '�������� ����������� ����

    Dim lpRgb As RGB, _
        lpFinalRgb As RGB
                    
    Dim c As Single
                
    Dim d As Integer, _
        e As Integer, _
        f As Integer
                                
    Dim x As Integer, _
        y As Integer, _
        z As Integer
        
    With lpGradient
        MkRGB .BaseColor, lpRgb.Red, lpRgb.Green, lpRgb.Blue
        MkRGB .FinalColor, lpFinalRgb.Red, lpFinalRgb.Green, lpFinalRgb.Blue
        c = .Percentage
    End With
    
    d = (CInt(lpFinalRgb.Red) - CInt(lpRgb.Red)) * c
    e = (CInt(lpFinalRgb.Green) - CInt(lpRgb.Green)) * c
    f = (CInt(lpFinalRgb.Blue) - CInt(lpRgb.Blue)) * c
    
    d = CInt(lpRgb.Red) + d
    e = CInt(lpRgb.Green) + e
    f = CInt(lpRgb.Blue) + f
        
    
    GetGradColor = RGB(d, e, f)

End Function

Private Sub MkRGB(ByVal ColorRef As Long, Red, Green, Blue) '���������� RGB
    Red = (ColorRef And &HFF)
    Green = ((ColorRef And &HFF00) / &H100) And &HFF
    Blue = ((ColorRef And &HFF0000) / &H10000) And &HFF
End Sub

Private Function AlphaBlendRGB(ByVal Color As Long, Color2 As Long, Number As Long, Highly As Boolean, Hg As Variant) As Long
Dim src1 As RGBTHREE
Dim src2 As RGBTHREE
Dim nRed As Long
Dim nGreen As Long
Dim nBlue As Long
   DoEvents
    
    GetRGB Color, src1
    GetRGB Color2, src2
If Highly = True Then
    DoEvents
    '����� �������������� ��������� � ������� ����� ������ �����
        nRed = (src1.rgbtRed * (50 + Hg * Number) + src2.rgbtRed * (100 - (50 + Hg * Number))) \ 100
        nGreen = (src1.rgbtGreen * (50 + Hg * Number) + src2.rgbtGreen * (100 - (50 + Hg * Number))) \ 100
        nBlue = (src1.rgbtBlue * (50 + Hg * Number) + src2.rgbtBlue * (100 - (50 + Hg * Number))) \ 100
        On Error Resume Next
        AlphaBlendRGB = RGB(nRed, nGreen, nBlue)
ElseIf Highly = False Then
    DoEvents
        nRed = (src1.rgbtRed * (50 + (50 - (Hg * Number))) + src2.rgbtRed * (100 - (50 + (50 - (Hg * Number))))) \ 100
        nGreen = (src1.rgbtGreen * (50 + (50 - (Hg * Number))) + src2.rgbtGreen * (100 - (50 + (50 - (Hg * Number))))) \ 100
        nBlue = (src1.rgbtBlue * (50 + (50 - (Hg * Number))) + src2.rgbtBlue * (100 - (50 + (50 - (Hg * Number))))) \ 100
        On Error Resume Next
        AlphaBlendRGB = RGB(nRed, nGreen, nBlue)
End If

End Function

Private Sub Massive() '��������� ��������� ������
Dim nX As Long, nY As Long
Dim SrcColor As Long
ReDim PixelRGB(0 To SizeX, 0 To SizeY) '�������������� ������

    For nX = 1 To SizeX / 2
        'DoEvents
        For nY = 1 To SizeY / 2
           SrcColor = AlphaBlendRGB(FinishColor, StartColor, nX, True, k)
           PixelRGB(nX, nY) = AlphaBlendRGB(SrcColor, StartColor, nY, True, H)
        Next nY
        For nY = SizeY / 2 To SizeY
           SrcColor = AlphaBlendRGB(FinishColor, StartColor, nX, True, k)
           PixelRGB(nX, nY) = AlphaBlendRGB(SrcColor, StartColor, nY - (SizeY / 2), False, H)
        Next nY
    Next nX
    For nX = SizeX / 2 To SizeX
        'DoEvents
        For nY = 1 To SizeY / 2
           SrcColor = AlphaBlendRGB(FinishColor, StartColor, nX - (SizeX / 2), False, k)
           PixelRGB(nX, nY) = AlphaBlendRGB(SrcColor, StartColor, nY, True, H)
        Next nY
        For nY = SizeY / 2 To SizeY
           SrcColor = AlphaBlendRGB(FinishColor, StartColor, nX - (SizeX / 2), False, k)
           PixelRGB(nX, nY) = AlphaBlendRGB(SrcColor, StartColor, nY - (SizeY / 2), False, H)
        Next nY
    Next nX
End Sub

Private Sub Gradient() '������ �� �������
Dim i As Long, j As Long, bits() As Long, bih As BITMAPINFOHEADER
With bih
    .biSize = Len(bih)
    .biWidth = SizeX
    .biHeight = SizeY
    .biBitCount = 32
    .biPlanes = 1
    .biSizeImage = 4 * .biWidth * .biHeight
    If .biSizeImage = 0 Then Exit Sub
    ReDim bits(0 To .biWidth - 1, 0 To .biHeight - 1)
End With

For i = 1 To SizeX
    For j = 1 To SizeY
        lpGradient.FinalColor = PixelRGB(i, j)
        lpBrush.lbColor = GetGradColor(lpGradient)
        bits(i - 1, j - 1) = lpBrush.lbColor
    Next j
Next i
pctBut.Cls
SetDIBitsToDevice pctBut.hdc, 0, 0, SizeX, SizeY, 0, 0, 0, SizeY, bits(0, 0), bih, 0
End Sub

Private Sub lblname_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    Timer1.Tag = "SomeDown"
Else
    Timer1.Tag = "Up"
End If
End Sub

Private Sub Timer1_Timer()
Dim R As Integer, G As Integer, B As Integer
DoEvents

Dim pnt As POINT_API
GetCursorPos pnt
ScreenToClient UserControl.hWnd, pnt
If pnt.x < UserControl.ScaleLeft Or _
       pnt.y < UserControl.ScaleTop Or _
       pnt.x > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
      Timer1.Tag = "Down"
Else
    Timer1.Tag = "Up"
End If

If Timer1.Tag = "Up" And lpGradient.Percentage < 0.9 Then
DoEvents
MkRGB lblName.ForeColor, R, G, B
If R - R1 > 0 Then
     R = R - 30
     lblName.ForeColor = RGB(R, G, B)
End If
If G - G1 > 0 Then
    G = G - 30
    lblName.ForeColor = RGB(R, G, B)
End If
If B - B1 > 10 Then
    B = B - 30
    lblName.ForeColor = RGB(R, G, B)
End If
lpGradient.Percentage = lpGradient.Percentage + 0.1
Gradient
ElseIf Timer1.Tag = "Down" And lpGradient.Percentage > 0.45 Then
DoEvents
MkRGB lblName.ForeColor, R, G, B
If R < 150 Then
    R = R + 30
    lblName.ForeColor = RGB(R, G, B)
End If
If G < 150 Then
    G = G + 30
    lblName.ForeColor = RGB(R, G, B)
End If
If B < 150 Then
    B = B + 30
    lblName.ForeColor = RGB(R, G, B)
End If

lpGradient.Percentage = lpGradient.Percentage - 0.07
Gradient
ElseIf lpGradient.Percentage = 0.45 Then
    lblName.ForeColor = RGB(150, 150, 255)
End If

End Sub

Private Sub UserControl_Initialize()
lpGradient.BaseColor = RGB(0, 0, 0)
StartColor = RGB(0, 0, 0)
FinishColor = RGB(255, 0, 0)
lpGradient.Percentage = 0.45

End Sub
Private Sub UserControl_Resize()
'--------------

pctBut.Left = 0
pctBut.Top = 0
pctBut.Width = UserControl.ScaleWidth
pctBut.Height = UserControl.ScaleHeight


'--------------------
lblName.Left = 0
lblName.Width = UserControl.ScaleWidth
lblName.Top = (UserControl.ScaleHeight - lblName.Height) / 2

SizeY = pctBut.ScaleHeight
SizeX = pctBut.ScaleWidth
MkRGB lblName.ForeColor, R1, G1, B1
k = 100 / SizeX ' �� �����������
H = 100 / SizeY

lpBrush.lbColor = GetGradColor(lpGradient)
hBrush = CreateBrushIndirect(lpBrush)

    Massive
    Gradient
'____________
End Sub
Private Function GetRGB(ByVal ColorVal As Long, Result As RGBTHREE) As Boolean
    Result.rgbtRed = ColorVal And 255
    Result.rgbtGreen = (ColorVal And 65535) \ 256
    Result.rgbtBlue = ColorVal \ 65536
End Function

Private Sub pctbut_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub
Private Sub pctbut_Click()
   RaiseEvent Click
End Sub
Private Sub pctbut_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub pctbut_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub pctbut_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub pctbut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub pctbut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Private Sub lblname_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub
Private Sub lblname_Click()
    RaiseEvent Click
End Sub
Private Sub lblname_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub lblname_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub lblname_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub lblname_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub lblname_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Sub StartButton()
Timer1.Enabled = True

End Sub

'�������� ����
Public Property Let FColor(ByVal vNewFinishColor As OLE_COLOR)
    FinishColor = vNewFinishColor
    Gradient
    PropertyChanged "FColor"
End Property
Public Property Get FColor() As OLE_COLOR
    FColor = FinishColor

End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "SColor", StartColor
    PropBag.WriteProperty "FColor", FinishColor
    PropBag.WriteProperty "BColor", lpGradient.BaseColor
    PropBag.WriteProperty "Caption", lblName.Caption
    PropBag.WriteProperty "Font", lblName.Font
    PropBag.WriteProperty "CColor", lblName.ForeColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag) '��������� �����������
    StartColor = PropBag.ReadProperty("SColor", StartColor)
    FinishColor = PropBag.ReadProperty("FColor", FinishColor)
    lpGradient.BaseColor = PropBag.ReadProperty("BColor", lpGradient.BaseColor)
    lblName.Caption = PropBag.ReadProperty("Caption", lblName.Caption)
    lblName.Font = PropBag.ReadProperty("Font", lblName.Font)
    lblName.ForeColor = PropBag.ReadProperty("CColor", lblName.ForeColor)
    Massive
    Gradient
    UserControl.Refresh
End Sub

Public Property Let SColor(ByVal vNewStartColor As OLE_COLOR)
    StartColor = vNewStartColor
    Gradient
    PropertyChanged "SColor"
End Property
Public Property Get SColor() As OLE_COLOR
    SColor = StartColor
    Gradient
End Property

Public Property Let BaseColor(ByVal vNewBaseColor As OLE_COLOR)
    lpGradient.BaseColor = vNewBaseColor
    Gradient
    PropertyChanged "BColor"
End Property
Public Property Get BaseColor() As OLE_COLOR
    BaseColor = lpGradient.BaseColor
End Property
Public Property Let Font(ByVal New_Font As StdFont)
Set lblName.Font = New_Font
Call UserControl_Initialize
PropertyChanged "Font"

End Property
Public Property Get Font() As StdFont
Set Font = lblName.Font
End Property

Public Property Let Caption(ByVal New_Caption As String)
lblName.Caption = New_Caption
Call UserControl_Initialize
PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
Caption = lblName.Caption
End Property

Public Property Let CaptionColor(ByVal vNewCapColor As OLE_COLOR)
    lblName.ForeColor = vNewCapColor
    Gradient
    PropertyChanged "CColor"
End Property
Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = lblName.ForeColor
End Property
