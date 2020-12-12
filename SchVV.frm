VERSION 5.00
Begin VB.Form SchVV 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2748
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   4584
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Проставить норматив"
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ок"
      DownPicture     =   "SchVV.frx":0000
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      WhatsThisHelpID =   3
      Width           =   1455
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
      Left            =   2640
      TabIndex        =   4
      Text            =   "0"
      Top             =   1200
      WhatsThisHelpID =   2
      Width           =   1815
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
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "SchVV.frx":078F
      Height          =   255
      Left            =   4320
      Picture         =   "SchVV.frx":0BAA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Закрыть"
      Top             =   0
      WhatsThisHelpID =   4
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
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
      Left            =   1080
      TabIndex        =   18
      Top             =   240
      Width           =   852
   End
   Begin VB.Label Label13 
      Caption         =   "Прописано"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1080
      TabIndex        =   17
      Top             =   0
      Width           =   732
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Расчет по нормативу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Left            =   2520
      TabIndex        =   16
      Top             =   360
      Width           =   1932
   End
   Begin VB.Label Label11 
      Caption         =   "Норматив:"
      Height          =   252
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   972
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
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
      TabIndex        =   14
      Top             =   240
      Width           =   852
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1920
      WhatsThisHelpID =   200
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1920
      WhatsThisHelpID =   200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1320
      TabIndex        =   9
      Top             =   1920
      WhatsThisHelpID =   200
      Width           =   372
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      WhatsThisHelpID =   200
      Width           =   972
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      WhatsThisHelpID =   200
      Width           =   1212
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      WhatsThisHelpID =   105
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Текущий период"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      WhatsThisHelpID =   102
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Предыдущий период"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      WhatsThisHelpID =   101
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ввод показаний счетчика"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      WhatsThisHelpID =   100
      Width           =   2652
   End
End
Attribute VB_Name = "SchVV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(Replace(Text2.Text, ".", ",")) < Val(Replace(Text1.Text, ".", ",")) Then
MsgBox ("Текущее значение счетчика не может быть, менее предыдущего")
'Exit Sub
End If

On Error GoTo tex
Lic.fg1.TextMatrix(Lic.fg1.Row, 42) = Replace(Text2.Text, ".", ",")
Lic.fg1.TextMatrix(Lic.fg1.Row, 41) = Replace(Text1.Text, ".", ",")

tex:
If Err.Number = 1004 Then
MsgBox ("Необходимо ввеси число")
Err.Clear
Exit Sub
End If
'Проверяем на норматив
If Me.Text2.Text = Lic.fg1.TextMatrix(Lic.fg1.Row, 41) + Lic.fg1.TextMatrix(Lic.fg1.Row, 47) * Lic.fg1.TextMatrix(Lic.fg1.Row, 12) Then
'Это норматив
Lic.fg1.TextMatrix(Lic.fg1.Row, 49) = True
Else
'Это НЕ норматив
Lic.fg1.TextMatrix(Lic.fg1.Row, 49) = False
End If
Unload Me
Sch.Refresh
End Sub

Private Sub Command2_Click()
Command1_Click
End Sub


Private Sub Command3_Click()
' Проставляем норматив
Me.Text2.Text = Lic.fg1.TextMatrix(Lic.fg1.Row, 41) + Lic.fg1.TextMatrix(Lic.fg1.Row, 47) * Lic.fg1.TextMatrix(Lic.fg1.Row, 12)
Lic.fg1.TextMatrix(Lic.fg1.Row, 49) = True

'Me.Text2.Text =
'Label8 = Round((Lic.fg1.TextMatrix(Lic.fg1.Row, 42) - Lic.fg1.TextMatrix(Lic.fg1.Row, 41)) * Lic.fg1.TextMatrix(Lic.fg1.Row, 10), 2)
Label8 = Lic.fg1.TextMatrix(Lic.fg1.Row, 42) - Lic.fg1.TextMatrix(Lic.fg1.Row, 41)

Label8.Refresh
End Sub

Private Sub Form_Load()
Text1.Text = Lic.fg1.TextMatrix(Lic.fg1.Row, 41)
Text2.Text = Lic.fg1.TextMatrix(Lic.fg1.Row, 42)
'Прописано
Label14.Caption = Lic.fg1.TextMatrix(Lic.fg1.Row, 12)


Label5.Caption = Lic.fg1.TextMatrix(Lic.fg1.Row, 41)
Label6.Caption = Lic.fg1.TextMatrix(Lic.fg1.Row, 42)
Label8 = Lic.fg1.TextMatrix(Lic.fg1.Row, 42) - Lic.fg1.TextMatrix(Lic.fg1.Row, 41)
'Норматив
Label10.Caption = Lic.fg1.TextMatrix(Lic.fg1.Row, 47)
If Lic.fg1.TextMatrix(Lic.fg1.Row, 49) = True Then
Label12.Visible = True
'Label12.Caption = Lic.fg1.TextMatrix(Lic.fg1.Row, 49)
Else
Label12.Visible = False
End If

End Sub

