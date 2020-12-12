VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl CalendarControl 
   BackColor       =   &H80000004&
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   4605
   ScaleWidth      =   3000
   ToolboxBitmap   =   "CalendarControl.ctx":0000
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2475
      Top             =   45
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2550
      Top             =   795
   End
   Begin VB.PictureBox PicSel 
      Height          =   420
      Left            =   1800
      Picture         =   "CalendarControl.ctx":0314
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   495
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox PicNow 
      Height          =   330
      Left            =   2340
      Picture         =   "CalendarControl.ctx":05EC
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   495
      Visible         =   0   'False
      Width           =   420
   End
   Begin XPControls.BlueCommand BlueCommand1 
      Height          =   375
      Left            =   1845
      TabIndex        =   1
      Top             =   0
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox EditableCalendar 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   -2147483644
      MaxLength       =   10
      Mask            =   "##.##.####"
      PromptChar      =   " "
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   135
      ScaleHeight     =   2895
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   1305
      Visible         =   0   'False
      Width           =   2600
      Begin VB.Shape Shape3 
         BorderColor     =   &H007F9DB9&
         Height          =   285
         Left            =   2295
         Top             =   0
         Width           =   285
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H007F9DB9&
         Height          =   285
         Left            =   0
         Top             =   0
         Width           =   285
      End
      Begin VB.Line Line1 
         X1              =   -90
         X2              =   2745
         Y1              =   675
         Y2              =   675
      End
   End
   Begin PicClip.PictureClip pc 
      Left            =   1575
      Top             =   1125
      _ExtentX        =   1588
      _ExtentY        =   476
      _Version        =   393216
      Cols            =   4
      Picture         =   "CalendarControl.ctx":0930
   End
   Begin MSMask.MaskEdBox mebTime 
      Height          =   300
      Left            =   1305
      TabIndex        =   5
      Top             =   45
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.Shape Shape1 
      Height          =   1185
      Left            =   0
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "CalendarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "KERNEL32" () As Long

Private Const LOCALE_SSHORTDATE = &H1F
Private Const LOCALE_SLONGDATE = &H20
Private Const ERROR_INSUFFICIENT_BUFFER = 122
Private Const ERROR_INVALID_FLAGS = 1004
Private Const ERROR_INVALID_PARAMETER = 87

Dim WithEvents cPrivate As clsPrivate
Attribute cPrivate.VB_VarHelpID = -1

Dim fMask As String
Dim frmt As String
Dim CurDate As Date
Dim lFull As Boolean
Dim arr(1 To 6, 1 To 7)

Public Event Change()

Public Property Set DataSource(rhs As ADODB.Recordset)
    Set EditableCalendar.DataSource = rhs
End Property

Public Property Let DataField(rhs As String)
    EditableCalendar.DataField = rhs
End Property

Public Property Get DataField() As String
    DataField = EditableCalendar.DataField
End Property

Public Property Let BackColor(bc As OLE_COLOR)
    EditableCalendar.BackColor = bc
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = EditableCalendar.BackColor
End Property

Public Property Let CustomFormat(ByVal NewCustomFormat As String)
    frmt = NewCustomFormat
End Property

Public Property Get CustomFormat() As String
    CustomFormat = frmt
End Property

Public Property Let FullTime(rhs As Boolean)
    lFull = rhs
    mebTime.Visible = rhs
End Property

Public Property Get FullTime() As Boolean
    FullTime = lFull
End Property

Public Property Let Value(ByVal NewDate As Variant)
    If IsNull(NewDate) Or IsEmpty(NewDate) Then
        ClearDate
    Else
        On Error GoTo ErrorHandler
        If Not IsEmpty(NewDate) Then CurDate = Format(Replace(NewDate, "T", " "), "dd.mm.yyyy")
        EditableCalendar.Text = Format(NewDate, frmt)
        mebTime.Text = Format(NewDate, "hh:mm")
    End If
    PropertyChanged "Value"
    Exit Property
ErrorHandler:
    ClearDate
    Resume
End Property

Private Function ValidateValueInernal() As Variant
Dim cd As Date
    If EditableCalendar.ClipText = "" Then
        ValidateValueInernal = Null
    Else
        On Error Resume Next
        CurDate = EditableCalendar.Text
        If Err.Number = 0 Then
            If mebTime.ClipText <> "" Then
                Err.Clear
                cd = mebTime.Text
                If Err.Number = 0 Then CurDate = CDate(Format(EditableCalendar.Text & " " & mebTime.Text, "dd.mm.yyyy hh:mm"))
            End If
            EditableCalendar.Text = CurDate
            ValidateValueInernal = CurDate
        Else
            ValidateValueInernal = Null
        End If
    End If
End Function

Public Property Get Value() As Variant
    Value = ValidateValueInernal
End Property

Public Property Get Text() As String
    Text = EditableCalendar.ClipText
End Property

Public Property Get Enabled() As Boolean
    Enabled = EditableCalendar.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    EditableCalendar.Enabled = vNewValue
    BlueCommand1.Enabled = vNewValue
    PropertyChanged ("Enabled")
End Property

Private Sub ClearDate()
    EditableCalendar.Mask = ""
    EditableCalendar.Text = ""
    mebTime.Mask = ""
    mebTime.Text = ""
    EditableCalendar.Refresh
    mebTime.Refresh
    CurDate = Date
End Sub

Private Sub make_xpbutton(i As Integer, z As Integer)
Dim brx, bry, bw, bh
Dim f
    On Error Resume Next
    f = Picture1.Font
    Picture1.Font = "Webdings"
    Picture1.FontSize = 10
    Picture1.ScaleMode = vbPixels
    bw = ScaleX(300, vbTwips, vbPixels) - 2
    bh = ScaleY(300, vbTwips, vbPixels) - 2
    If i = 1 Then
        Picture1.PaintPicture pc.GraphicCell(z), 1, 1, bw, bh, 1, 1, 13, 16
        Picture1.CurrentX = 3
        Picture1.CurrentY = 1
        Picture1.Print "3"
    Else
        brx = ScaleX(Picture1.Width - 300, vbTwips, vbPixels)
        Picture1.PaintPicture pc.GraphicCell(z), brx, 1, bw, bh, 1, 1, 13, 16
        Picture1.CurrentX = brx + 2
        Picture1.CurrentY = 1
        Picture1.Print "4"
    End If
    Picture1.ScaleMode = vbTwips
    Picture1.Font = f
End Sub

Private Function MakeMask(ByVal DateFormat As String) As String
Dim result As Variant
Dim string_index As Long
    If Len(DateFormat) = 0 Then
        MakeMask = ""
        Exit Function
    End If
    frmt = DateFormat
    DateFormat = Replace(DateFormat, "y", "#", 1, , vbTextCompare)
    DateFormat = Replace(DateFormat, "h", "#", 1, , vbTextCompare)
    DateFormat = Replace(DateFormat, "s", "#", 1, , vbTextCompare)
    result = InStr(1, "mmm", DateFormat, vbTextCompare)
    If result <> 0 Then
        string_index = result
        While (StrComp(Mid$(DateFormat, string_index, 1), "m") = 0)
            DateFormat = Replace(DateFormat, "m", string_index, 1, vbTextCompare)
            string_index = string_index + 1
        Wend
    Else
        DateFormat = Replace(DateFormat, "m", "#", 1, , vbTextCompare)
    End If
    result = InStr(1, "ddd", DateFormat, vbTextCompare)
    If result <> 0 Then
        string_index = result
        While (StrComp(Mid$(DateFormat, string_index, 1), "d") = 0)
            DateFormat = Replace(DateFormat, "d", string_index, 1, vbTextCompare)
            string_index = string_index + 1
        Wend
    Else
        DateFormat = Replace(DateFormat, "d", "#", 1, , vbTextCompare)
    End If
    If InStr(GetDateFormat, "/") <> 0 Then
        MakeMask = Replace(DateFormat, ".", "/")
    Else
        MakeMask = DateFormat
    End If
End Function

Private Sub BlueCommand1_Click()
Dim cp As POINTAPI
    If Not Enabled Then Exit Sub
    On Error GoTo errh
    ClientToScreen UserControl.hwnd, cp
    UserControl_Show
    CreateDD ObjPtr(cPrivate), Picture1.hwnd, cp.X, cp.Y + ScaleY(UserControl.Height, vbTwips, vbPixels), ScaleX(Picture1.Width, vbTwips, vbPixels), ScaleY(Picture1.Height, vbTwips, vbPixels)
    'RaiseEvent Click
    Exit Sub
errh:
End Sub

Private Sub EditableCalendar_Change()
    PropertyChanged "Value"
End Sub

Private Sub EditableCalendar_GotFocus()
    If EditableCalendar.Mask = "" Then EditableCalendar.Mask = MakeMask(frmt)
    EditableCalendar.SelStart = 0
    EditableCalendar.SelLength = Len(EditableCalendar.Text)
End Sub

Private Sub EditableCalendar_LostFocus()
Dim valDate
    valDate = ValidateValueInernal
    If IsNull(valDate) Then ClearDate
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 330 Then
        If X < 330 Then
            make_xpbutton 1, 2
        ElseIf X > Picture1.Width - 330 Then
            make_xpbutton 2, 2
        End If
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 330 Then
        If X < 330 Then
            make_xpbutton 1, 1
            Timer1.Enabled = True
            cPrivate.sHideOnClick = False
        ElseIf X > Picture1.Width - 330 Then
            make_xpbutton 2, 1
            Timer2.Enabled = True
            cPrivate.sHideOnClick = False
        Else
            cPrivate.sHideOnClick = True
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, j As Integer
    If Y < 330 Then
        If X < 330 Then
            make_xpbutton 1, 1
            CurDate = DateAdd("m", -1, CurDate)
            UserControl_Show
        ElseIf X > Picture1.Width - 330 Then
            make_xpbutton 2, 1
            CurDate = DateAdd("m", 1, CurDate)
            UserControl_Show
        End If
    Else
        For j = 1 To 7
            For i = 1 To 6
                If X < (j + 1) * 350 - 250 And Y < (i + 1) * 300 + 500 Then
                    CurDate = arr(i, j)
                    UserControl_Show
                    Value = CurDate
                    RaiseEvent Change
                    Exit Sub
                End If
            Next i
        Next j
    End If
End Sub

Private Sub Timer1_Timer()
Dim pnt As POINTAPI
    GetCursorPos pnt
    ScreenToClient Picture1.hwnd, pnt
    Picture1.ScaleMode = vbPixels
    If pnt.X < Picture1.ScaleLeft Or pnt.Y < Picture1.ScaleTop Or _
        pnt.X > (Picture1.ScaleLeft + ScaleX(330, vbTwips, vbPixels)) Or _
        pnt.Y > (Picture1.ScaleTop + ScaleY(330, vbTwips, vbPixels)) Then
            Timer1.Enabled = False
            Picture1.ScaleMode = vbTwips
            make_xpbutton 1, 0
    End If
    Picture1.ScaleMode = vbTwips
End Sub

Private Sub Timer2_Timer()
Dim pnt As POINTAPI
    GetCursorPos pnt
    ScreenToClient Picture1.hwnd, pnt
    Picture1.ScaleMode = vbPixels
    If pnt.X < (Picture1.ScaleLeft + Picture1.ScaleWidth - ScaleX(330, vbTwips, vbPixels)) Or _
        pnt.Y < Picture1.ScaleTop Or pnt.X > (Picture1.ScaleLeft + Picture1.ScaleWidth) Or _
        pnt.Y > (Picture1.ScaleTop + ScaleY(330, vbTwips, vbPixels)) Then
            Timer2.Enabled = False
            Picture1.ScaleMode = vbTwips
            make_xpbutton 2, 0
    End If
    Picture1.ScaleMode = vbTwips
End Sub

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    EditableCalendar.SetFocus
End Sub

Private Sub UserControl_Initialize()
    EditableCalendar.Mask = ""
    EditableCalendar.Text = ""
    CurDate = Date
    Shape1.BorderColor = RGB(127, 157, 185)
    Set cPrivate = New clsPrivate
    Picture1.BackColor = vbWhite
    'cPrivate.sHideOnClick = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim date_format As String
Dim cf As String
    cf = PropBag.ReadProperty("CustomFormat", "")
    If InStr(GetDateFormat, "/") <> 0 Then
        CustomFormat = Replace(cf, ".", "/")
    Else
        CustomFormat = cf
    End If
    DataField = PropBag.ReadProperty("DataField", "")
    date_format = GetDateFormat
    EditableCalendar.Mask = MakeMask(date_format)
    BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    FullTime = PropBag.ReadProperty("FullTime", False)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    BlueCommand1.Left = UserControl.Width - BlueCommand1.Width
    BlueCommand1.Height = UserControl.Height
    mebTime.Left = BlueCommand1.Left - mebTime.Width
    mebTime.Height = UserControl.Height - 75
    EditableCalendar.Height = UserControl.Height - 75
    If lFull Then
        EditableCalendar.Width = UserControl.Width - BlueCommand1.Width - mebTime.Width - 60
    Else
        EditableCalendar.Width = UserControl.Width - BlueCommand1.Width - 60
    End If
    Shape1.Width = UserControl.Width
    Shape1.Height = UserControl.Height
End Sub

Private Sub UserControl_Show()
Dim i As Integer, j As Integer
Dim X As Integer
Dim fd As Date
    With Picture1
        .Cls
        Picture1.Line (0, 0)-(.Width, 0), RGB(172, 168, 153)
        Picture1.Line (0, 0)-(0, .Height), RGB(172, 168, 153)
        Picture1.Line (.Width - 15, 0)-(.Width - 15, .Height), RGB(172, 168, 153)
        Picture1.Line (0, .Height - 15)-(.Width, .Height - 15), RGB(172, 168, 153)
        .FontBold = True
        .CurrentY = 70
        .CurrentX = (.Width - TextWidth(MonthName(Month(CurDate)) & " " & Year(CurDate))) / 2 - 150
        Picture1.Print MonthName(Month(CurDate)) & " " & Year(CurDate)
        For X = 1 To 7
            .CurrentY = 400
            .CurrentX = X * 350 - 250
            Picture1.Print WeekdayName(X, True, vbMonday)
        Next X
        fd = DateAdd("d", -(Day(CurDate) - 1), CurDate)
        i = Weekday(fd, vbMonday)
        X = 1
        For j = i - 1 To 1 Step -1
            arr(1, j) = DateAdd("d", -X, fd)
            X = X + 1
        Next j
        X = 0
        For j = i To 7
            arr(1, j) = DateAdd("d", X, fd)
            X = X + 1
        Next j
        For i = 2 To 6
            For j = 1 To 7
                arr(i, j) = DateAdd("d", X, fd)
                X = X + 1
            Next j
        Next i
        .FontBold = False
        For j = 1 To 7
            For i = 1 To 6
                .CurrentX = j * 350 - 250
                .CurrentY = i * 300 + 500
                If CurDate = arr(i, j) Then .PaintPicture PicSel.Picture, .CurrentX, .CurrentY, 320, 220, 0, 0, 300, 200
                If Date = arr(i, j) Then .PaintPicture PicNow.Picture, .CurrentX, .CurrentY, 320, 220, 0, 0, 300, 200
                If Month(arr(i, j)) = Month(CurDate) Then
                    .ForeColor = vbBlack
                    .FontBold = False
                Else
                    .ForeColor = RGB(172, 168, 153)
                    .FontBold = False
                End If
                Picture1.Print Day(arr(i, j))
            Next i
        Next j
        .CurrentY = 2600
        .ForeColor = vbBlack
        .FontBold = True
        .CurrentX = (.Width - TextWidth("Сегодня " & Format(Date, "dd.mm.yyyy"))) / 2
        Picture1.Print "Сегодня " & Format(Date, "dd.mm.yyyy")
        make_xpbutton 1, 0
        make_xpbutton 2, 0
    End With
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DataField", DataField, ""
    PropBag.WriteProperty "CustomFormat", CustomFormat
    PropBag.WriteProperty "BackColor", EditableCalendar.BackColor
    PropBag.WriteProperty "FullTime", FullTime
End Sub

Public Property Let MaskFormat(rhs As String)
    fMask = rhs
End Property

Private Function GetDateFormat() As String
Dim result As Long
Dim date_format As String
Dim LocaleID As Long
    LocaleID = GetUserDefaultLCID()
    result = GetLocaleInfo(LocaleID, LOCALE_SSHORTDATE, date_format, 0)
    If result <> ERROR_INSUFFICIENT_BUFFER And _
        result <> ERROR_INVALID_FLAGS And _
        result <> ERROR_INVALID_PARAMETER Then
        date_format = Space(result - 1) 'надо чего-нибудь в переменную положить, иначе после вызова GetLocaleInfo возвращается пустая строка
        result = GetLocaleInfo(LocaleID, LOCALE_SSHORTDATE, date_format, result)
    End If
    GetDateFormat = date_format
End Function


