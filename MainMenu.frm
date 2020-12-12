VERSION 5.00
Begin VB.Form MainMenu 
   BackColor       =   &H80000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5364
   ClientLeft      =   12
   ClientTop       =   228
   ClientWidth     =   10164
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   847
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command15 
      BackColor       =   &H000000FF&
      Caption         =   "Срок ГО истек. Как продлить ГО"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   10092
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   252
      Left            =   7440
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Command14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4110
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Просмотр данных архива"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Default         =   -1  'True
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Закрыть период F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6720
      Picture         =   "MainMenu.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command10 
      Height          =   495
      Left            =   9600
      Picture         =   "MainMenu.frx":1447
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Который час? "
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Документы оплаты F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "MainMenu.frx":3141
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Настройки F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6720
      MaskColor       =   &H80000013&
      Picture         =   "MainMenu.frx":34F6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ОДН F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      MaskColor       =   &H80000013&
      Picture         =   "MainMenu.frx":387A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Расчет F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MaskColor       =   &H80000018&
      Picture         =   "MainMenu.frx":39B3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Тарифы F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MaskColor       =   &H80000013&
      Picture         =   "MainMenu.frx":3DAC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Отчеты. Анализ F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      MaskColor       =   &H00FFFFC0&
      Picture         =   "MainMenu.frx":426F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Caption         =   "Выход F12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      MaskColor       =   &H80000013&
      Picture         =   "MainMenu.frx":43A7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Справочники F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      MaskColor       =   &H80000013&
      Picture         =   "MainMenu.frx":47E9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.Image imgTitleHelp 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      Height          =   156
      Left            =   0
      Picture         =   "MainMenu.frx":4C0A
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   156
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   """Квартплата +""  (C) Copyright, 2005, Астрахань, Бугоров Андрей Владимирович. Консультации +79881733-600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   9855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "123"
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10176
   End
   Begin VB.Image imgTitleMain 
      Height          =   360
      Left            =   1920
      Picture         =   "MainMenu.frx":4E54
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   288
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   1080
      Picture         =   "MainMenu.frx":559E
      Top             =   360
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   720
      Picture         =   "MainMenu.frx":5CE8
      Top             =   360
      Width           =   228
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Расчет 
         Caption         =   "Расчет"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Тарифы 
         Caption         =   "Тарифы"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Справ 
         Caption         =   "Справ"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Документы 
         Caption         =   "Документы"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Улицы 
         Caption         =   "Улицы"
         Shortcut        =   {F6}
      End
      Begin VB.Menu Отчеты 
         Caption         =   "Отчеты"
         Shortcut        =   {F7}
      End
      Begin VB.Menu Настройки 
         Caption         =   "Настройки"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Период 
         Caption         =   "Период"
         Shortcut        =   {F9}
      End
      Begin VB.Menu ArhivFalse 
         Caption         =   "Отменить запреты на архив"
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu Активировать 
         Caption         =   "Активировать"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu Выход 
         Caption         =   "Выход"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Sprav.Show , MainForm
MainMenu.Hide
End Sub

Private Sub Command10_Click()
'DataReport1.Show
Closc.Show
End Sub

Private Sub Command11_Click()
Zakr.Show
End Sub

Private Sub Command12_Click()
GoTo en
' Это по <Т> первым символам, а хотелось бы покруче проверить,
' т.е. просто убрать сразу "-", или если буква пропущена?

Dim t As Integer 'Точность проверки
Dim strA As String 'Строка в которой ищем
Dim strB As String 'Строка в котоую ищем ПРИБЛИЗИТЕЛЬНАЯ

t = 5

strA = "Иванова"
strB = "Иван-ов"

For I = 1 To Len(strB) - t + 1
MsgBox strB

If InStr(1, strA, strB) Then MsgBox "Равно" Else MsgBox "Нет"
strB = Mid(strB, 1, Len(strB) - 1)

Next
en:
strA = "Иванова"
strB = "Петров"


MsgBox Compare(strA, strB, 10)

End Sub

Private Sub Command13_1_Click()

End Sub

Private Sub Command13_Click()
Unload Me
ArhivDialog.Show
End Sub


Private Sub Command14_Click()
'Количество lyrq в месяц
a = DateDiff("d", MainForm.DR, DateAdd("m", 1, MainForm.DR))
'MsgBox (a)
MainForm.DnP = DateDiff("d", MainForm.DR, DateAdd("m", 1, MainForm.DR))
  
'MsgBox (Month(MainForm.DR))
End Sub

Private Sub Command15_Click()
Form_GO.Label1.Visible = False
Form_GO.Show 1
End Sub

Private Sub Command2_Click()
MainMenu.Hide
MainForm.Hide
'Call BaseProtect(App.Path + "\data\kvartplata.amd", True)
End
End Sub

Private Sub Command3_Click()
'Form1.Show
MainMenu.Enabled = False
Reports.Show
End Sub

Private Sub Command4_Click()
'MainMenu.Hide
Me.Enabled = False
Tarif.Show
End Sub

Private Sub Command5_Click()
'TMP.Show


MainMenu.Enabled = False
Pass.Show

End Sub

Private Sub Command6_Click()
Command6.BackColor = &H80000010
Command6.Refresh
Command6.Caption = "Пожалуйста подождите"
MainMenu.Enabled = False
Filter.Show
Filter.SetFocus
MainMenu.Hide
End Sub

Private Sub Command7_Click()
ODN_MENU.Show


'****************
MainMenu.Hide
End Sub

Private Sub Command8_Click()
MainMenu.Enabled = False
MenuNastr.Show

End Sub

Private Sub Command9_Click()
'ReestrDoc.Show
MenuDoc.Show
MainMenu.Enabled = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If (KeyCode = 113 And MainForm.strDataName <> "kvartplata.amd") Then
 '     Arhiv = False
    
  ' End If




If KeyCode = 113 Then
      Call Command6_Click
   End If
If KeyCode = 114 Then
      Call Command4_Click
   End If

If KeyCode = 115 Then
      Call Command1_Click
   End If
If KeyCode = 116 Then
      Call Command9_Click
   End If
   If KeyCode = 117 Then
      Call Command7_Click
   End If
   If KeyCode = 118 Then
      Call Command3_Click
   End If
   
   If KeyCode = 119 Then
      Call Command8_Click
   End If
   If KeyCode = 123 Then
      Call Command2_Click
   End If
End Sub

Private Sub Form_Load()


Me.KeyPreview = True

Menu.Visible = False

MakeWindow Me, True


'Command5.Enabled = False
If MainForm.strDataName <> "kvartplata.amd" And Arhiv = True Then
Arhiv = True
lblTitle.Caption = "Расчет коммунальных платежей населения АРХИВ"
lblTitle.ForeColor = vbRed
Me.Command1.BackColor = &HC0FFC0
Me.Command2.BackColor = &HC0FFC0
Me.Command3.BackColor = &HC0FFC0
Me.Command4.BackColor = &HC0FFC0
Me.Command5.BackColor = &HC0FFC0
Me.Command6.BackColor = &HC0FFC0
Me.Command7.BackColor = &HC0FFC0
Me.Command8.BackColor = &HC0FFC0
Me.Command9.BackColor = &HC0FFC0
Me.Command10.BackColor = &HC0FFC0
Me.Command11.BackColor = &HC0FFC0
Me.Command11.Enabled = False
Me.Command1.Enabled = False
Me.Command7.Enabled = False
Me.Command4.Enabled = False
Else
Arhiv = False
lblTitle.Caption = "Расчет коммунальных платежей населения " + MainForm.Label7 + "     Кол-во лиц/сч >" + Str(MainForm.LcKol)

lblTitle.ToolTipText = "Кол-во лиц/сч >" + Str(MainForm.LcKol) + ". В т.ч. договоров-" + Str(MainForm.LcKolD) + ". Абон.книжек-" + Str(MainForm.LcKolK)
'Label1.ForeColor = &H8000000D
Me.Command13.Caption = MainForm.Label8

End If
End Sub

Private Sub imgTitleHelp_Click()
'About.Show

Dim AboutBox As New AboutBox
With AboutBox
    .Title = " Расчет и анализ коммунальных платежей населения"
    .Version = "Версия: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision)
    .Company = "Квартплата +  (C) Copyright, 2005, Астрахань"
    .Copyright = " Бугоров Андрей Владимирович"
    .Description = "Комплексная автоматизация бухучета"
    .License = "Связь с автором E-Mail:bestonline@list.ru телефоны:+79881733600"
    .hWndOwner = Me.hwnd
    'Set .Icon = Me.Icon
    .AboutBox
End With

End Sub

Private Sub imgTitleHelp_DblClick()
About.Show
End Sub

Private Sub Активировать_Click()
MainForm.Enabled = True
End Sub

Private Sub Выход_Click()
Command2_Click
End Sub

Private Sub Документы_Click()
Command9_Click
End Sub

Private Sub Настройки_Click()
Command8_Click
End Sub

Private Sub Отчеты_Click()
Command3_Click
End Sub

Private Sub Период_Click()
'Command11_Click
End Sub

Private Sub Расчет_Click()
Command6_Click
End Sub

Private Sub Справ_Click()
Command1_Click
End Sub

Private Sub Тарифы_Click()
Command4_Click
End Sub

Private Sub Улицы_Click()
Command7_Click
End Sub
