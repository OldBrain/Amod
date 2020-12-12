VERSION 5.00
Begin VB.Form Comentariy 
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   885
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "_"
      Top             =   120
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   8520
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   8520
      X2              =   8520
      Y1              =   840
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   8510
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Comentariy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text1.Text) <= 200 Then
Lic.FG1.TextMatrix(Lic.FG1.Row, 38) = Text1.Text
Else
MsgBox ("Слишком длинная строка комментария")
Exit Sub
End If

Lic.Enabled = True
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
MsgBox (Str(KeyAscii))
End Sub

Private Sub Form_Load()
Lic.Enabled = False

Text1.Text = Lic.FG1.TextMatrix(Lic.FG1.Row, 38)


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Lic.Enabled = True
Unload Me
End If
End Sub
