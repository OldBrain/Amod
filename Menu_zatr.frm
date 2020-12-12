VERSION 5.00
Begin VB.Form Menu_zatr 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3096
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   8796
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1092
   End
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8532
      _ExtentX        =   15050
      _ExtentY        =   656
      Caption         =   "Анализ затрат"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   8535
      _ExtentX        =   15050
      _ExtentY        =   656
      Caption         =   "Закрыть"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8535
      _ExtentX        =   15050
      _ExtentY        =   656
      Caption         =   "1. Настройка "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   192
      Left            =   0
      Picture         =   "Menu_zatr.frx":0000
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   192
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   960
      Picture         =   "Menu_zatr.frx":0542
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
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "123"
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   8250
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1320
      Picture         =   "Menu_zatr.frx":0C8C
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   480
      Picture         =   "Menu_zatr.frx":13D6
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "Menu_zatr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show

End Sub

Private Sub Form_Load()
MakeWindow Me, True
lblTitle.Caption = "Расчет и анализ затрат на содержание ЖКХ"

End Sub

Private Sub imgTitleHelp_Click()
xpcmdbutton2_Click
End Sub

Private Sub xpcmdbutton1_Click()
Z_Nastr.Show
Me.Hide
End Sub

Private Sub xpcmdbutton2_Click()
MainMenu.Show
Unload Me
End Sub

Private Sub xpcmdbutton3_Click()
Anal_Zatrat.Show
Me.Hide
End Sub
