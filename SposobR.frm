VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SposobR 
   BorderStyle     =   0  'None
   ClientHeight    =   4536
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7488
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   Icon            =   "SposobR.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "SposobR.frx":030A
   ScaleHeight     =   4536
   ScaleWidth      =   7488
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   7200
      Picture         =   "SposobR.frx":2D112
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      MaskColor       =   &H000000FF&
      Picture         =   "SposobR.frx":2D520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   7215
      _ExtentX        =   12721
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "С перерасчетом льгот (занимает на много больше времени)"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   120
      Picture         =   "SposobR.frx":2DAB3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4800
      Picture         =   "SposobR.frx":2DEB3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2640
      Picture         =   "SposobR.frx":2E422
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   480
      Picture         =   "SposobR.frx":2EB9B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "С перерасчетом льгот (занимает на много больше времени)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Способ расчета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Расчет"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "SposobR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Dolgo As Boolean
Public RasRr As Boolean

Private Sub Command1_Click()
Picture1.Visible = False
Label4.Visible = False
Label3.Caption = "Ждите идет расчет начислений"
RasRr = False
Command5.Visible = True
Command5.Enabled = True

Command5.SetFocus

Check1.Visible = False
ProgressBar1.Visible = True

Dolgo = False
If Check1.Value = 1 Then Dolgo = True Else Dolgo = False

'Unload SposobR
SposobR.Enabled = False
Filter.Enabled = False


Label1 = "Пожалуйста подождите. Идет расчет."
'SposobR.Hide

Расчет1.Isprav = 1 'С учетом исправлений
Расчет1.Show
Расчет1.Visible = False
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
Label4.Visible = False
Label3.Caption = "Ждите идет расчет начислений"

RasRr = False
Command5.Visible = True
Command5.Enabled = True
Command5.SetFocus

Check1.Visible = False
ProgressBar1.Visible = True

Label1 = "Пожалуйста подождите. Идет расчет."
'Unload SposobR
SposobR.Enabled = False
Filter.Enabled = False
Расчет1.Isprav = 2 'Без учета исправлений
Расчет1.Show
Расчет1.Visible = False
End Sub

Private Sub Command3_Click()
Filter.Enabled = True
Unload Me
End Sub

Private Sub Command4_Click()
Unload Расчет1
Unload Me
End Sub

Private Sub Command5_Click()

If MsgBox("Прервать расчет?", vbYesNo) = vbYes Then
Unload Filter
End
'RasRr = True
'Filter.Enabled = True
'Unload Me
'Unload Расчет1
'Else

End If

End Sub

Private Sub Form_Load()
Filter.Enabled = False
ProgressBar1.Visible = False
ProgressBar1.min = 0
ProgressBar1.Max = Filter.Fg.Rows


End Sub

Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True
End Sub

Private Sub Picture1_Click()
Command3_Click
End Sub
