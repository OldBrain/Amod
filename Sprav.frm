VERSION 5.00
Begin VB.Form Sprav 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4092
   ClientLeft      =   12
   ClientTop       =   216
   ClientWidth     =   9984
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Sprav.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   832
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник затрат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   3372
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник домов"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   3372
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник типов льгот <F8>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник счетов затрат <F7>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник видов расчетов <F6>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник типов квартир <F5>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник типов домов <F4>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник соцминимумов <F3>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник категорий расчетов<F2>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      Picture         =   "Sprav.frx":030A
      ScaleHeight     =   2244
      ScaleWidth      =   2244
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0CCCC&
         Caption         =   """Квартплата+"" "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0CCC1&
      Caption         =   "В Ы Х О Д <F10>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0CCC1&
      Caption         =   "Справочник льгот <F1>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   3375
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
      Picture         =   "Sprav.frx":3793
      Top             =   0
      Width           =   156
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
      Left            =   240
      TabIndex        =   10
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10050
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "Sprav.frx":39DD
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "Sprav.frx":4127
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   360
      Picture         =   "Sprav.frx":4871
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Льготы 
         Caption         =   "Льготы"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Sprav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Lgota.Show
Sprav.Hide
End Sub

Private Sub Command10_Click()
Doma.Show
End Sub

Private Sub Command11_Click()
Menu_zatr.Show
Sprav.Hide
End Sub

Private Sub Command2_Click()
MainMenu.Show
Sprav.Hide
End Sub

Private Sub Command3_Click()
Kategor.Show
End Sub

Private Sub Command4_Click()
Socmin.Show
Sprav.Hide
End Sub

Private Sub Command5_Click()
TipDom.Show
Sprav.Hide
End Sub

Private Sub Command6_Click()
TipKv.Show
Sprav.Hide
End Sub

Private Sub Command7_Click()
Nachisleniy.Show
Sprav.Hide
Sprav.Enabled = False
End Sub

Private Sub Command8_Click()
Schet1.Show 1, Me
'Sprav.Enabled = False
End Sub

Private Sub Command9_Click()
lgtip.Show
Sprav.Enabled = False
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_Load()
Меню.Visible = False
MakeWindow Me, False
lblTitle.Caption = "Заполнение справочников"
Label2.BackColor = RGB(207, 207, 207)

Me.Label2.Caption = Label2.Caption + vbNewLine + MainForm.Label7

'Me.Caption = "Заполнение справочников"
End Sub

Private Sub imgTitleHelp_Click()
About.Show
End Sub

Private Sub Льготы_Click()
Lgota.Show
Sprav.Enabled = False
End Sub
