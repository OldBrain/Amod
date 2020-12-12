VERSION 5.00
Begin VB.Form PopUp 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   6336
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11508
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Квартиросъемщик"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   2172
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   580
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      X1              =   0
      X2              =   800
      Y1              =   50
      Y2              =   50
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000D&
      Caption         =   "предприятие:"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   5876
      TabIndex        =   8
      Top             =   360
      Width           =   1092
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "Сегодня:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   852
   End
   Begin VB.Label PeriodR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   11520
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   " Текущий расчетный период :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   252
      Left            =   1906
      TabIndex        =   5
      Top             =   360
      Width           =   2772
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   1068
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Демо версия программы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   2616
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7920
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4332
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   252
      Left            =   4671
      TabIndex        =   1
      Top             =   360
      Width           =   1212
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
      Picture         =   "PopUp.frx":0000
      ToolTipText     =   "О программе"
      Top             =   0
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Квартплата + "
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
      Height          =   228
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1320
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   5040
      Picture         =   "PopUp.frx":0542
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   -1320
      Picture         =   "PopUp.frx":0C8C
      Top             =   360
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   360
      Left            =   3120
      Picture         =   "PopUp.frx":13D6
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   288
   End
End
Attribute VB_Name = "PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Filter.FG.SetFocus
End Sub

Private Sub Form_Load()
MakeWindow Me, True
Line1.X2 = Me.hwnd

End Sub

Private Sub Label6_Click()

End Sub
