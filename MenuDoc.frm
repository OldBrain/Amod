VERSION 5.00
Begin VB.Form MenuDoc 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2496
   ClientLeft      =   12
   ClientTop       =   120
   ClientWidth     =   6516
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Оплата/субсидии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Настраиваемые документы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   6495
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Данные счетчика"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   6495
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
      Picture         =   "MenuDoc.frx":0000
      Top             =   0
      Width           =   156
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "MenuDoc.frx":024A
      Top             =   0
      Width           =   228
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
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   5850
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "MenuDoc.frx":0994
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   360
      Picture         =   "MenuDoc.frx":10DE
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "MenuDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEnh1_1_Click()

End Sub

Private Sub BtnEnh1_Click()

ReestrDoc.Show
MainMenu.Enabled = False
Unload Me
End Sub

Private Sub BtnEnh2_Click()
'MsgBox "В стадии разработки"
MainMenu.Enabled = True
ReestrTablDoc.Show
Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub BtnEnh3_Click()
MainMenu.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub BtnEnh4_Click()

End Sub

Private Sub Command1_Click()
Unload Me
Sch_kat.Show

End Sub

Private Sub Form_Load()
MainMenu.Enabled = False

lblTitle = "Документы"
MakeWindow Me, True
End Sub

Private Sub imgTitleHelp_Click()
Dim AboutBox As New AboutBox
With AboutBox
    .Title = " Расчет и анализ коммунальных платежей населения"
    .Version = "Версия: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision)
    .Company = "Квартплата +  (C) Copyright, 2005, Астрахань"
    .Copyright = " Бугоров Андрей Владимирович"
    .Description = "Комплексная автоматизация бухучета"
    .License = "Связь с автором E-Mail:bestonline@list.ru телефоны: +7988-733-600"
    .hWndOwner = Me.hwnd
    'Set .Icon = Me.Icon
    .AboutBox
End With
End Sub

Private Sub lblTitle_Click()
About.Show

End Sub
