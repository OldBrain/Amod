VERSION 5.00
Begin VB.Form PoctaVibor 
   Caption         =   "Form5"
   ClientHeight    =   1620
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4644
   LinkTopic       =   "Form5"
   ScaleHeight     =   1620
   ScaleWidth      =   4644
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   4572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4572
   End
End
Attribute VB_Name = "PoctaVibor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public code As Integer

Private Sub Command1_Click()
code = Val(Combo1.Text)
Me.Hide
End Sub

Private Sub Form_Load()
Dim CombRs As ADODB.Recordset
Set CombRs = New ADODB.Recordset

CombRs.Open ("SELECT nachisleniy.Kod, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy WHERE (((nachisleniy.Tip)='-'))"), Mconn



' Заполняем Комбо
CombRs.MoveFirst
Do While Not CombRs.EOF
Combo1.AddItem CStr(CombRs("Kod")) & "  " & CombRs("Naim")
CombRs.MoveNext
Loop
CombRs.Close

Combo1.Text = "Введите код оплаты"


End Sub

