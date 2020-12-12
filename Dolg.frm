VERSION 5.00
Begin VB.Form FormDolg 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2220
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   6036
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Caption         =   "Отмена <Esc>"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Печать <Enter>"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   5775
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
      Picture         =   "Dolg.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   156
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Сумма долга"
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4650
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   360
      Picture         =   "Dolg.frx":024A
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "Dolg.frx":0994
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "Dolg.frx":10DE
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "FormDolg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo obratno
Lic.Dolg = Replace(Me.Text1, ".", ",")
obratno:
If Err.Number <> 0 Then
MsgBox (Err.Description)
Err.Clear
Text1.ForeColor = vbRed
Text1.SetFocus
Exit Sub
End If
'Unload Me
'If Lic.Dolg <= 0 Then
'Label1.Visible = True
'Label1.Caption = "Нулевая или отрицательная сумма долга!" + vbNewLine + " печать извещения не имеет сьысла"
'Text1.ForeColor = vbRed
'Text1.SetFocus
'Exit Sub
'End If
Unload Me
End Sub

Private Sub Command2_Click()
Lic.Dolg = -369.8985231
Unload Me
End Sub

Private Sub Form_Load()
MakeWindow Me, True

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Lic.Dolg = Me.Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
If KeyAscii = 27 Then Command2_Click
End Sub
