VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Doc1 
   AutoRedraw      =   -1  'True
   Caption         =   "Документ"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12090
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Разнести"
      Height          =   735
      Left            =   9840
      Picture         =   "Doc1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5040
      TabIndex        =   8
      Text            =   "Начисление"
      Top             =   840
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Адрес"
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   10455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   11295
      _cx             =   19923
      _cy             =   9975
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Doc1.frx":0102
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   2
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4440
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Doc1.frx":029E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Doc1.frx":03B0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Doc1.frx":04C2
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Расчитать"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label8 
      Caption         =   "Кол-во строк документа"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "И того:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "# ##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "# ##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Line Line9 
      X1              =   7080
      X2              =   7080
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line8 
      X1              =   960
      X2              =   960
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   120
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line5 
      X1              =   10680
      X2              =   10680
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line4 
      X1              =   10680
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   5640
      X2              =   5640
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line1 
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   10680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label4 
      Caption         =   "Начисление:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Адрес:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Начисление"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Адрес"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Добавить 
         Caption         =   "Добавить"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Удалить"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Закрыть 
         Caption         =   "Закрыть"
         Shortcut        =   {F12}
      End
      Begin VB.Menu Разнести 
         Caption         =   "Разнести по лиц.счета"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "Doc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_Adding, Rs, Rs_Combo, Rs_Combo1, Rs_Combo2, Cmb, CMB1 As ADODB.Recordset
Dim TheConn As ADODB.Connection
Dim sq1, Kod As String
Dim s, Nm, Nd, SS As Double
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Private Sub Combo2_Validate(Cancel As Boolean)
Label2 = Combo2.Text
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
Label1 = Combo3.Text
Rs_Combo1.Close
КомбоФИО

End Sub

Private Sub Command1_Click()
Doc.Enabled = False
Pod.Show

Разнести_Click
End Sub

Private Sub FG_AfterDataRefresh()
Цвет
FG.ColComboList(13) = "..."
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
ReestrDoc.FG.TextMatrix(ReestrDoc.FG.Row, 4) = Text1
End Sub

' Проверить ввод
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            Добавить_Click
        Case "Delete"
            Удалить_Click
        Case "Save"
            Закрыть_Click
    End Select
End Sub


Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'nn = FG.TextMatrix(FG.Row, 5)
'nk = 0
'For Rw = 1 To FG.Rows - 1
'NNN = FG.TextMatrix(Rw, 5)
'If NNN = nn And nn <> 0 Then nk = nk + 1
'If nk > 1 Then
'MsgBox ("Нельзя вносить два одинаковых номера л/сч в один документ! Для №" + nn + " " + FG.TextMatrix(Rw, 6) + " Прийдется создать новый документ!")
'FG.TextMatrix(Rw, 5) = 0
'FG.TextMatrix(Rw, 6) = "....."
'End If
'Next Rw



If FG.TextMatrix(FG.Row, FG.Col) = "" Then Exit Sub

Rs_Combo.MoveFirst
Do While Not Rs_Combo.EOF
         If Rs_Combo("Kod") = FG.TextMatrix(FG.Row, 3) Then
    FG.TextMatrix(FG.Row, 4) = Rs_Combo("Naim")
    FG.TextMatrix(FG.Row, 12) = Rs_Combo("Tip")
                  End If
Rs_Combo.MoveNext
Loop
'FG.TextMatrix(FG.Row, 5) = Val(FG.TextMatrix(FG.Row, 6))
b = InStr(1, FG.TextMatrix(FG.Row, 6), " ")
If b > 1 Then
A = Trim(Left(FG.TextMatrix(FG.Row, 6), b - 1))
Q = "SELECT MainOccupant.Numer, MainOccupant.FAM From MainOccupant WHERE(((MainOccupant.FAM)=" + Chr(34) + A + Chr(34) + "))"
Rs_Combo2.Open (Q)
On Error Resume Next
FG.TextMatrix(FG.Row, 5) = Rs_Combo2.Fields("Numer").Value
Rs_Combo2.Close
End If
If FG.TextMatrix(FG.Row, 5) = "" Then FG.TextMatrix(FG.Row, 5) = 0
Q = "SELECT MainOccupant.Numer, MainOccupant.FAM From MainOccupant WHERE(((MainOccupant.Numer)=" + FG.TextMatrix(FG.Row, 5) + "))"
'MsgBox (Q)
Rs_Combo2.Open (Q)
If Rs_Combo2.EOF = False Then FG.TextMatrix(FG.Row, 6) = Rs_Combo2.Fields("FAM").Value
Rs_Combo2.Close
'Rs.Redraw
FG.ComboList = ""

If FG.TextMatrix(FG.Row, 7) <> SS Then
'MsgBox (FG.TextMatrix(FG.Row, 8))
FG.TextMatrix(FG.Row, 10) = 1
TheConn.Execute ("UPDATE Doc INNER JOIN Adding ON Doc.Key = Adding.KodDoc SET Adding.SummaI =" + Str(FG.TextMatrix(FG.Row, 7)) + ", Adding.ispr = 1 WHERE (((Adding.KodDoc)=" + FG.TextMatrix(FG.Row, 8) + "))")
'MsgBox (Str(SS) + FG.TextMatrix(FG.Row, 10))
End If
Цвет
TheConn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
SS = FG.TextMatrix(FG.Row, 7)
If FG.TextMatrix(0, FG.Col) = "Код" Or FG.TextMatrix(0, FG.Col) = "Ф.И.О." Or FG.TextMatrix(0, FG.Col) = "№ л/сч" Then
       

' Начисления
If FG.TextMatrix(0, FG.Col) = "Код" Then
                            If Combo2.Text = "Любое начисление" Then
'FG.Editable = flexEDKbdMouse
cl = ""
Rs_Combo.MoveFirst
Do While Not Rs_Combo.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
cl = cl + CStr(Rs_Combo("Kod")) & vbTab & Rs_Combo("Naim") + "|"
Rs_Combo.MoveNext
Loop
FG.ComboList = cl
'Else: FG.ComboList = ""

'Else: FG.Editable = flexEDNone
                             End If

 End If
'Фамилии
If FG.TextMatrix(0, FG.Col) = "Ф.И.О." Then
'FG.Editable = flexEDKbdMouse
cl = ""
'If Not Rs_Combo1.EOF Then
Rs_Combo1.MoveFirst
j = 0
Do While Not Rs_Combo1.EOF
j = j + 1
If j > 1000 Then
MsgBox ("Слишком длинный для отоброжения список лицевах счетов(более 1000) Проставте пожалуйста адрес в шапке документа.")
Exit Do
End If
'cl = cl + Combo_RS("Name_Kategor") + "|"
'cl = cl + CStr(Rs_Combo1("Numer")) & vbTab & CStr(Rs_Combo1("ФИО")) & vbTab & Rs_Combo1("АДР") + "|"
'On Error Resume Next
If Rs_Combo1("ФИО") <> "" Then cl = cl + CStr(Rs_Combo1("ФИО")) & vbTab & CStr(Rs_Combo1("Numer")) & vbTab & Rs_Combo1("АДР") + "|"
Rs_Combo1.MoveNext
Loop
FG.ComboList = cl
'Else: FG.ComboList = ""
End If

'№ л/сч
If FG.TextMatrix(0, FG.Col) = "№ л/сч" And FG.TextMatrix(FG.Row, FG.Col) = "0" Then
'FG.Editable = flexEDKbdMouse
cl = ""
о = 0
'If Not Rs_Combo1.EOF Then
Rs_Combo1.MoveFirst
Do While Not Rs_Combo1.EOF
j = j + 1
If j > 1000 Then
MsgBox ("Слишком длинный для отоброжения список лицевах счетов(более 1000) Проставте пожалуйста адрес в шапке документа.")
Exit Sub
End If

'cl = cl + Combo_RS("Name_Kategor") + "|"
'If Rs_Combo1("ФИО") = nul Then Rs_Combo1("ФИО") = "......"
'On Error Resume Next
If Rs_Combo1("ФИО") <> "" Then cl = cl + CStr(Rs_Combo1("Numer")) & vbTab & CStr(Rs_Combo1("ФИО")) & vbTab & Rs_Combo1("АДР") + "|"
'cl = cl + CStr(Rs_Combo1("ФИО")) & vbTab & CStr(Rs_Combo1("Numer")) & vbTab & Rs_Combo1("АДР") + "|"
Rs_Combo1.MoveNext
Loop
FG.ComboList = cl
'Else: FG.ComboList = ""
End If
Else
FG.ComboList = ""
'FG.Editable = flexEDNone
End If
Итог
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

' Запрет на редоктирование ячеек в зависимости от номера колонки
If FG.Col = 1 Or FG.Col = 2 Or FG.Col = 4 Then Cancel = True
'MsgBox (FG.TextMatrix(0, FG.Col) + "  " + FG.TextMatrix(FG.Row, FG.Col))
If (FG.TextMatrix(0, FG.Col) = "Ф.И.О.") Then
If (FG.TextMatrix(FG.Row, FG.Col) <> ".........") Then Cancel = True
End If


If (FG.TextMatrix(0, FG.Col) = "№ л/сч") Then
If (FG.TextMatrix(FG.Row, FG.Col) <> "0") Then Cancel = True
End If


'Or FG.Col > 7
'If DgT.ColDataType(Col) = flexDTDate Then
End Sub

Private Sub Form_Load()
Set TheConn = New ADODB.Connection
TheConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
TheConn.Open "data/Kvartplata.mdb"
'Set rs_Tit = New ADODB.Recordset
   




'Recordset для фильтров
Set Cmb = New ADODB.Recordset
Set Cmb.ActiveConnection = TheConn
Set CMB1 = New ADODB.Recordset
Set CMB1.ActiveConnection = TheConn
Cmb.CursorType = adOpenForwardOnly
Cmb.LockType = adLockBatchOptimistic
Cmb.Open "Nachisleniy"
CMB1.Open "KLS_PODR"

' Заполняем Combo2 для начисления
'Set Combo2.DataSource = Combo_RS
Combo2.Text = ReestrDoc.FG.TextMatrix(ReestrDoc.FG.Row, 3)
cl = "Любое начисление"
Cmb.MoveFirst
Do While Not Cmb.EOF
Combo2.AddItem cl
cl = CStr(Cmb("Kod")) & "  " & Cmb("Naim")
'codN(Combo_RS("Kod")) = Combo_RS("Kod")
Cmb.MoveNext
Loop

' Заполняем Combo3 для адресов
'Set Combo2.DataSource = Cmb1

Combo3.Text = ReestrDoc.FG.TextMatrix(ReestrDoc.FG.Row, 10)
'cl = "0   Все дома  0"
CMB1.MoveFirst
Do While Not CMB1.EOF
If CMB1("Код") <> 0 Then
cl = CStr(CMB1("Код")) & "  " & CMB1("Naim_kls") & " дом № " & CMB1("Num")
Combo3.AddItem cl
End If
CMB1.MoveNext
Loop



Set Rs = New ADODB.Recordset
Set Rs.ActiveConnection = TheConn


Set Rs_Combo = New ADODB.Recordset
Set Rs_Combo.ActiveConnection = TheConn

Set Rs_Combo1 = New ADODB.Recordset
Set Rs_Combo1.ActiveConnection = TheConn

Set Rs_Combo2 = New ADODB.Recordset
Set Rs_Combo2.ActiveConnection = TheConn


Doc.Caption = "Документ на начисление/удержание/субсидию на дату " + ReestrDoc.FG.TextMatrix(ReestrDoc.FG.Row, 2)


FG.Editable = flexEDKbdMouse

Label1 = ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)
Label2 = ReestrDoc.FG.TextMatrix(ReestrDoc.r, 3)
Text1 = ReestrDoc.FG.TextMatrix(ReestrDoc.r, 4)
'rs_Tit.Open

Rs.CursorType = adOpenForwardOnly
Rs.LockType = adLockBatchOptimistic

Rs_Combo.CursorType = adOpenForwardOnly
Rs_Combo.LockType = adLockBatchOptimistic

Rs_Combo2.CursorType = adOpenForwardOnly
Rs_Combo2.LockType = adLockBatchOptimistic

Kod = ReestrDoc.FG.TextMatrix(ReestrDoc.r, 1)

Rs.Open ("SELECT Doc.*, Doc.Cod From Doc WHERE (((Doc.Cod)=" + Kod + "))")
Rs_Combo.Open "Nachisleniy"




 'Это выбор Recordset для Combo фамилий, взависимости от выбранного
 'адреса в шапке документа
  
sq1 = "SELECT MainOccupant.Numer, MainOccupant!FAM+" & Chr(34) & " " & Chr(34) + "+Left(MainOccupant!IM,1)+" + Chr(34) + ". " + Chr(34) + " AS ФИО, " & Chr(34) & "ул." & Chr(34) & "+KLS_PODR!NAIM_KLS+" & Chr(34)
sq1 = sq1 + "дом № " & Chr(34) & "+KLS_PODR!Num AS АДР, MainOccupant.Dom FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom=KLS_PODR.КОД"
kod1 = ""
'If Val(ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)) <> 0 Then
If Val(Combo3.Text) <> 0 Then
'Kod = Str(Val(ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)))
kod1 = Str(Val(Combo3.Text))
sq1 = sq1 + " WHERE (((MainOccupant.Dom)=" + kod1 + "))"
End If

Rs_Combo1.Open (sq1)









FG.DataMode = flexDMBoundImmediate
' Cвойства, свойства необходимые для сортировки
'    FG.AllowUserResizing = flexResizeBoth
 '   FG.ExtendLastCol = True
    FG.ExplorerBar = flexExSort
    FG.AutoSearch = flexSearchFromCursor


Set FG.DataSource = Rs

FG.ColComboList(13) = "..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
'ReestrDoc.Show
'ReestrDoc.FG.Refresh
'ReestrDoc.FG.Redraw
'Load ReestrDoc
End Sub

Private Sub Добавить_Click()

Rs.AddNew
Rs("doc.Cod") = ReestrDoc.FG.TextMatrix(ReestrDoc.r, 1)
Rs("DataR") = ReestrDoc.FG.TextMatrix(ReestrDoc.r, 2)
Rs("NameKv") = "........."

If Combo2.Text <> "Любое начисление" Then
Rs("KodN") = Val(Combo2.Text)

'ReestrDoc.FG.TextMatrix(ReestrDoc.R, 8)
Else
Rs("KodN") = 0
End If

Rs.UpdateBatch
FG.DataRefresh
Rs.MoveLast

Rs_Combo.MoveFirst
Do While Not Rs_Combo.EOF
If Rs_Combo("Kod") = FG.TextMatrix(FG.Row, 3) Then FG.TextMatrix(FG.Row, 4) = Rs_Combo("Naim")
Rs_Combo.MoveNext
Loop
If FG.TextMatrix(FG.Row, 5) = "" Then FG.TextMatrix(FG.Row, 5) = 0
TheConn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")


End Sub

Private Sub Закрыть_Click()
Unload ReestrDoc
'MsgBox (Label5)
'Unload ReestrDoc
'st = Str(Int(s)) + "," + Str(s - Int(s))
On Error Resume Next
kod1 = FG.TextMatrix(1, 1)
St = Doc.Label5
ad = Chr(34) + Combo3.Text + Chr(34)

If kod1 <> "" Then
TheConn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Summa = " + St + ",ReestrDoc.Adres = " + ad + ",ReestrDoc.coment = " + Chr(34) + Text1 + Chr(34) + "  WHERE (((ReestrDoc!Cod)=" + kod1 + "))")
End If
'ReestrDoc.FG.TextMatrix(ReestrDoc.r, 5) = s
'ReestrDoc.FG.TextMatrix(ReestrDoc.r, 4) = Doc.Label5
ReestrDoc.Hide
Unload Doc
Load ReestrDoc
ReestrDoc.Show
ReestrDoc.FG.DataRefresh
ReestrDoc.Refresh


End Sub

Private Sub Разнести_Click()
'Cmb.Close
'CMB1.Close
'Rs.Close
'Rs_Combo.Close


Set rs_Adding = New ADODB.Recordset
Set rs_Adding.ActiveConnection = TheConn
rs_Adding.CursorType = adOpenForwardOnly
rs_Adding.LockType = adLockBatchOptimistic

'rs_Adding.Open ("SELECT Adding.KodKv FROM Adding INNER JOIN Doc ON Adding.KodDoc = Doc.Key GROUP BY Adding.KodKv")





' Удаляем старые

Rs.MoveFirst
                                Do While Not Rs.EOF
N = Rs.Fields("Key").Value

TheConn.Execute ("DELETE Adding.*, Adding.KodDoc From Adding WHERE (((Adding.KodDoc)=" + Str(N) + "))")
Rs.MoveNext
                                       Loop
                                       
Nd = FG.TextMatrix(1, 1)

MsgBox (Nd)
Qdoc = "INSERT INTO Adding ( NameN, KodKat, KodN, KodKv, KodDoc, NameKat, DataR, Socmin, Propis, Projiv, ProLift, ObPl, PolPl, Formula, Tarif, Com, FLOOR, SchetZ, TarifD, TarifI, ispr, TipDomKod, TipKvKod, Tip, SummaI, Parametr, Lig, LgotaVid ) SELECT nachisleniy.Naim, nachisleniy.КодKategor, Doc.KodN, Doc.KodKv, Doc.Key, nachisleniy.Kategor, Doc.DataR, Socmin.Value, MainOccupant.NLODGERF, MainOccupant.NLODGER, MainOccupant.NLODLIFT, MainOccupant.COMSPACE, MainOccupant.HABSPACE, nachisleniy.Formula, Tarif.Value, Doc.Com, MainOccupant.FLOOR, nachisleniy.SchetZ, Tarif.TarifD, Tarif.TarifI, Doc.Stst, Tarif.KodDOM, Tarif.KodKV, nachisleniy.Tip, Doc.Summa, " + Chr(34) + "Не определено" + Chr(34) + " AS Выражение1, nachisleniy.Lig, nachisleniy.Vid "


TheConn.Execute (Qdoc + "FROM (Socmin INNER JOIN (MainOccupant INNER JOIN (nachisleniy INNER JOIN Doc ON nachisleniy.Kod = Doc.KodN) ON MainOccupant.Numer = Doc.KodKv) ON (Socmin.koli = MainOccupant.NLODGERF) AND (Socmin.Kategor = nachisleniy.Kategor)) INNER JOIN Tarif ON (MainOccupant.KV = Tarif.KodKV) AND (nachisleniy.КодKategor = Tarif.KodKat) AND (MainOccupant.DomTip = Tarif.KodDOM) WHERE (((Doc.Cod)=" + Nd + "))")
                                       
                    'Добавляем новые
'1.Если небыло исправлений вручную
'TheConn.Execute ("INSERT INTO Adding ( KodKv, DataR, KodN, NameN, KodDoc, ispr, com, tip ) SELECT Doc.KodKv, Doc.DataR, Doc.KodN, Doc.NameN, Doc.Key, 0, doc.com, doc.tip  From Doc WHERE (((Doc.Cod)=" + Kod + ") and (doc.stst=0))")
'2.Если были исправления вручную
'TheConn.Execute ("INSERT INTO Adding ( KodKv, DataR, KodN, NameN, KodDoc, SummaI, ispr, com, tip ) SELECT Doc.KodKv, Doc.DataR, Doc.KodN, Doc.NameN, Doc.Key, Doc.Summa, doc.stst, doc.com, doc.tip From Doc WHERE (((Doc.Cod)=" + Kod + ") and (doc.stst=1))")

'Заполняем остальные пустые поля Adding
'If Not rs_Adding.EOF Then rs_Adding.MoveFirst
                               'Do While Not rs_Adding.EOF
                             '  For Rw = 1 To FG.Rows - 1
                             ' Nm = FG.TextMatrix(Rw, 5)
                              'Nd = FG.TextMatrix(Rw, 1)
                              
                              
'                             Обновить



                            'rs_Adding.MoveNext
                                  'Next Rw
'MsgBox (FG.TextMatrix(1, 1))
TheConn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 1 WHERE (((ReestrDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")

'If MsgBox("Записи из документа разнесены. Пересчитать лицевые счета квартиросъемщиков, входящих в данный документ", vbYesNo) = vbYes Then
Pod.Label1.Caption = "Данные разнесены успешно"
Pod.Command1.Visible = True

'SposobR2.Show
'End If

End Sub

Private Sub Удалить_Click()
'TheConn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")
Dim DelItem As String
With Rs
DelItem = FG.TextMatrix(FG.Row, 8)

If MsgBox("Вы хотите удалить начисление " + FG.TextMatrix(FG.Row, 4) + " для " + FG.TextMatrix(FG.Row, 6) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If Rs("Key") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
FG.DataRefresh
'MsgBox (DelItem)
TheConn.Execute ("DELETE Adding.KodDoc, Adding.* From Adding WHERE (((Adding.KodDoc)= " + DelItem + "))")
TheConn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")
If .EOF Then .MoveLast
End If
End With

End Sub

Private Sub КомбоФИО()
 'Это выбор Recordset для Combo фамилий, взависимости от выбранного
 'адреса в шапке документа
 
sq1 = "SELECT MainOccupant.Numer, MainOccupant!FAM+" & Chr(34) & " " & Chr(34) + "+Left(MainOccupant!IM,1)+" + Chr(34) + ". " + Chr(34) + " AS ФИО, " & Chr(34) & "ул." & Chr(34) & "+KLS_PODR!NAIM_KLS+" & Chr(34)
sq1 = sq1 + "дом № " & Chr(34) & "+KLS_PODR!Num AS АДР, MainOccupant.Dom FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom=KLS_PODR.КОД"
Kod = ""
'If Val(ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)) <> 0 Then
If Val(Combo3.Text) <> 0 Then
'Kod = Str(Val(ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)))
Kod = Str(Val(Combo3.Text))
sq1 = sq1 + " WHERE (((MainOccupant.Dom)=" + Kod + "))"
End If

Rs_Combo1.Open (sq1)


End Sub
Private Sub Итог()
Dim s As Double
Dim Kol As Integer

s = 0
Kol = 0
For Rw = 1 To FG.Rows - 1
s = s + Round(FG.TextMatrix(Rw, 7), 2)
Kol = Kol + 1
Next Rw
Label5 = Str(s)
Label6 = Str(Kol)
End Sub
Private Sub Обновить()
ComboQ = "Where(((Adding.KodKv) = " & Nm & "))"
ComboQ = "Where(((Adding.KodKv) = " & Nm & " and (adding.Koddoc)= " + Nd + "))"
'TheConn.Execute ("UPDATE Adding INNER JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.NameN = [Nachisleniy]![Naim], Adding.KodKat = [Nachisleniy]![КодKategor], Adding.Formula = [Nachisleniy]![Formula], Adding.Tip = [Nachisleniy]![Tip], Adding.NameKat = [Nachisleniy]![Kategor] " + ComboQ)
' Дата расчета
'TheConn.Execute ("UPDATE Settings, Adding SET  Adding.DataR = [Settings]![TekData]" + ComboQ)
'Прочие

'TheConn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.TipKvKod = [MainOccupant]![KV], Adding.TipDomKod = [MainOccupant]![DomTip]" + СomboQ)
'Соцминимум
'TheConn.Execute ("UPDATE Adding SET Adding.Socmin =0 " + ComboQ)
'TheConn.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.Propis = Socmin.koli) AND (Adding.KodKat = Socmin.KodKategor) SET Adding.Socmin = [Socmin]![Value]" + ComboQ)
'Тариф
'TheConn.Execute ("UPDATE Adding SET Adding.Tarif = 0 " + ComboQ)
'TheConn.Execute ("UPDATE Adding INNER JOIN Tarif ON (Tarif.KodDOM = Adding.TipDomKod) AND (Tarif.KodKV = Adding.TipKvKod) AND (Adding.KodKat = Tarif.KodKat) SET Adding.Tarif = [Tarif]![Value]" + ComboQ)
'Сальдо
'Заполнить статус ИСПРАВЛЕНО 0 если небыло исправлений вручную
'TheConn.Execute ("UPDATE Adding SET Adding.ispr = 0 WHERE (((Adding.ispr)<>1) and ((Adding.KodKv) = " & Nm & ") and ((adding.Koddoc)= " + Nd + "))")

TheConn.Execute ("UPDATE Adding LEFT JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.LgotaVid = [Nachisleniy]![Vid]" + ComboQ)

'Rs.Requery
End Sub

Private Sub Цвет()
Dim Rw As Integer
For Rw = 1 To FG.Rows - 1
'MsgBox (fg1.TextMatrix(fw, 27))
If FG.TextMatrix(Rw, 10) = 1 Then
'FG1.Cell(flexcpFontBold, Rw, 1, Rw, FG1.Cols) = True
'fg1.Cell(flexcpBackColor, Rw, 0) = vbCyan

FG.Cell(flexcpFontBold, Rw, 7, Rw, 7) = True
FG.Cell(flexcpBackColor, Rw, 7, Rw, 7) = vbCyan
End If
'If fg1.TextMatrix(Rw, 23) = "+" Then fg1.Cell(flexcpForeColor, Rw, 18, Rw, 18) = vbBlue
'If fg1.TextMatrix(Rw, 23) = "-" Then fg1.Cell(flexcpForeColor, Rw, 18, Rw, 18) = vbRed
Next Rw
'fg1.Refresh

End Sub

Private Sub fg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI
    MsgBox ("111")
    ' get popup window position
    
    pt.X = FG.ColPos(Col) / Screen.TwipsPerPixelX
    pt.Y = (FG.RowPos(Row) + FG.RowHeight(Row)) / Screen.TwipsPerPixelY
    ClientToScreen FG.hwnd, pt
    
    ' show date popup
    'If fg.ColDataType(Col) = flexDTDate Then
    If FG.TextMatrix(FG.Row, FG.Col) = "....." Then
    
        With frmDate
        
            '.lblRow = Row
            '.lblCol = Col
            .Tag = FG.Cell(flexcpText, Row, Col)
            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
            .Show
        End With
        Exit Sub
    End If
    
    ' show file popup
    If InStr(FG.Cell(flexcpText, 0, Col), "File") > 0 Then
        With frmFile
            .lblRow = Row
            .lblCol = Col
            .Tag = FG.Cell(flexcpText, Row, Col)
            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
            .Show
        End With
        Exit Sub
    End If

End Sub

