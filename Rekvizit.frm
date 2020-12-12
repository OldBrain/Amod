VERSION 5.00
Begin VB.Form Rekvizit 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5760
   ClientLeft      =   2730
   ClientTop       =   3330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   FillColor       =   &H00800000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   13
      Text            =   "0"
      Top             =   720
      Width           =   5775
   End
   Begin VB.TextBox Text5 
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
      Left            =   120
      TabIndex        =   11
      Text            =   "0"
      Top             =   4560
      Width           =   5895
   End
   Begin VB.TextBox Text4 
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
      Left            =   120
      TabIndex        =   9
      Text            =   "0"
      Top             =   3720
      Width           =   5895
   End
   Begin VB.TextBox Text3 
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
      Left            =   2640
      TabIndex        =   7
      Text            =   "0"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text2 
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
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   1800
      Width           =   5775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Полное наименование предприятия"
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
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Реквизиты банка"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
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
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   5730
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
      Height          =   195
      Left            =   0
      Picture         =   "Rekvizit.frx":0000
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   0
      Picture         =   "Rekvizit.frx":024A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   720
      Picture         =   "Rekvizit.frx":0994
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1080
      Picture         =   "Rekvizit.frx":10DE
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Расчетный счет"
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
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   5775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Кор.счет"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "ИНН"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "БИК"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Банк"
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
      Top             =   1440
      Width           =   5535
   End
End
Attribute VB_Name = "Rekvizit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsSet As ADODB.Recordset

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
MakeWindow Me, True
MainForm.RsSet.Requery
Me.Text1 = MainForm.Bank
Me.Text2 = MainForm.BIK
Me.Text3 = MainForm.INN
Me.Text4 = MainForm.KS
Me.Text5 = MainForm.RS
Me.Text6 = MainForm.NamePr


Label1.BackColor = RGB(207, 207, 207)
Label2.BackColor = RGB(207, 207, 207)
Label3.BackColor = RGB(207, 207, 207)
Label4.BackColor = RGB(207, 207, 207)
Label5.BackColor = RGB(207, 207, 207)
Label6.BackColor = RGB(207, 207, 207)
CancelButton.BackColor = &H80000016
OKButton.BackColor = &H80000016
End Sub

Private Sub OKButton_Click()

'Set RsSet = New ADODB.Recordset
'RsSet.Open ("setting"), Mconn

MainForm.RsSet.MoveFirst
MainForm.RsSet.Fields("Bank").Value = Trim(Me.Text1.Text)
MainForm.Bank = Me.Text1

MainForm.RsSet.Fields("BIK").Value = Trim(Me.Text2.Text)
 MainForm.BIK = Me.Text2
 
MainForm.RsSet.Fields("INN").Value = Trim(Me.Text3.Text)
 MainForm.INN = Me.Text3
 
MainForm.RsSet.Fields("KS").Value = Trim(Me.Text4.Text)
 MainForm.KS = Me.Text4
 
MainForm.RsSet.Fields("RS").Value = Trim(Me.Text5.Text)
 MainForm.RS = Me.Text5
 
 MainForm.RsSet.Fields("name").Value = Trim(Me.Text6.Text)
 MainForm.NamePr = Me.Text6
 
MainForm.RsSet.UpdateBatch
'MainForm.RsSet.Requery
Unload Me
End Sub

Private Sub Text2_LostFocus()
If Len(Trim(Text2)) <> 9 Then
MsgBox ("Длина счета не равна 9 символов")
Text2.ForeColor = vbRed
Else
Text2.ForeColor = vbBlack
End If
End Sub

Private Sub Text4_LostFocus()
If Len(Trim(Text4)) <> 20 Then
MsgBox ("Длина счета не равна 20 символов")
Text4.ForeColor = vbRed
Else
Text4.ForeColor = vbBlack
End If
End Sub

Private Sub Text5_LostFocus()
If Len(Trim(Text5)) <> 20 Then
MsgBox ("Длина счета не равна 20 символов")
Text5.ForeColor = vbRed
Else
Text5.ForeColor = vbBlack
End If
End Sub
