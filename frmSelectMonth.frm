VERSION 5.00
Begin VB.Form frmSelectMonth 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1875
   ClientLeft      =   15
   ClientTop       =   -60
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelectMonth.frx":0000
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   4200
      Picture         =   "frmSelectMonth.frx":2D1F
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Декабрь"
      Height          =   315
      Index           =   11
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Ноябрь"
      Height          =   315
      Index           =   10
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Октябрь"
      Height          =   315
      Index           =   9
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Сентябрь"
      Height          =   315
      Index           =   8
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Август"
      Height          =   315
      Index           =   7
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Июль"
      Height          =   315
      Index           =   6
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Июнь"
      Height          =   315
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Май"
      Height          =   315
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Апрель"
      Height          =   315
      Index           =   3
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Март"
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Февраль"
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Январь"
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Дата"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Укажите месяц и год"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmSelectMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lReturnYear As Long
Public lReturnMonth As Long
Public sReturnMonth As String
Public Ok As Boolean

Private Sub cmdMonth_Click(Index As Integer)
lReturnYear = txtYear
lReturnMonth = Index + 1
sReturnMonth = cmdMonth(Index).Caption
Ok = True
'MsgBox cmdMonth(Index).Caption
'MsgBox Index
Doc.FG.TextMatrix(Doc.FG.Row, 15) = CDate("01." + Str(lReturnMonth) + "." + Str(lReturnYear))

Unload Me
'Me.Hide
End Sub

Private Sub cmdNext_Click()
txtYear = txtYear + 1
End Sub

Private Sub cmdPrevious_Click()
txtYear = txtYear - 1
End Sub

Private Sub Form_Load()
Me.Label2.Caption = Doc.FG.TextMatrix(Doc.FG.Row, 15)
'txtYear = Year(Date)
'cmdMonth(Month(Doc.Fg.TextMatrix(Doc.Fg.Row, 15))).Style = 1
cmdMonth(Month(Doc.FG.TextMatrix(Doc.FG.Row, 15)) - 1).TabIndex = 0
cmdMonth(Month(Doc.FG.TextMatrix(Doc.FG.Row, 15)) - 1).BackColor = vbYellow


'vbBlue
'vbMagenta
'vbRed
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub

Private Sub txtYear_Validate(Cancel As Boolean)
If IsNumeric(txtYear) = False Then Cancel = True
End Sub
