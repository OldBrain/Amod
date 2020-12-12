VERSION 5.00
Begin VB.Form Form_GO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2664
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8196
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   HasDC           =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   133.2
   ScaleMode       =   2  'Point
   ScaleWidth      =   409.8
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   $"GO.frx":0000
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   8172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   8052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Для продления срога гарантийного обслуживания программы пожалуйста обратитесь к разработчику по телефону: 8(8512)433-600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   852
      Left            =   0
      TabIndex        =   4
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   8172
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
      Picture         =   "GO.frx":00A5
      ToolTipText     =   "О программе"
      Top             =   0
      Visible         =   0   'False
      Width           =   156
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Срок ГО подходит к концу"
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
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "123"
      Top             =   1800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   3336
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   4440
      Picture         =   "GO.frx":02EF
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   6120
      Picture         =   "GO.frx":0A39
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   360
      Left            =   5280
      Picture         =   "GO.frx":1183
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   168
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "До завершения срока гарантийного обслуживания программы осталось"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   8052
   End
End
Attribute VB_Name = "Form_GO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
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

Private Sub Form_Load()
MakeWindow Me, False
End Sub

Private Sub Label3_Click()

End Sub

