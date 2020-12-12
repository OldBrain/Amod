VERSION 5.00
Begin VB.Form Dni1 
   Caption         =   "Введите количество дней"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4296
   LinkTopic       =   "Form4"
   ScaleHeight     =   2400
   ScaleWidth      =   4296
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Количество дней"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2652
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1212
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4212
   End
End
Attribute VB_Name = "Dni1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'MsgBox (Lic.fg1.TextMatrix(Lic.fg1.Row, 22))

'Mconn.Execute ("UPDATE Adding SET Adding.DnF = " + Me.Text1.Text + " WHERE (((Adding.KodKv)=" + MainForm.Fnum + "))")
Mconn.Execute ("UPDATE Adding SET Adding.DnF = " + Me.Text1.Text + " WHERE (((Adding.KodKv)=" + MainForm.Fnum + ") AND ((Adding.KodKat)=" + Lic.fg1.TextMatrix(Lic.fg1.Row, 22) + "))")


Lic.fg1.TextMatrix(Lic.fg1.Row, 46) = Me.Text1.Text

Lic.Command16.Caption = "Дни" + "-" + Lic.fg1.TextMatrix(Lic.fg1.Row, 45) + "/" + Lic.fg1.TextMatrix(Lic.fg1.Row, 46)
'Me.Text1.Text
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = Lic.Caption
Text1.Text = Lic.fg1.TextMatrix(Lic.fg1.Row, 46)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Lic.Enabled = True
End Sub
