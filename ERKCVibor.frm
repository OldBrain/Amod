VERSION 5.00
Begin VB.Form ERKCVibor 
   Caption         =   "Выбор кода оплаты"
   ClientHeight    =   1776
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   4992
   LinkTopic       =   "Form3"
   ScaleHeight     =   1776
   ScaleWidth      =   4992
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4572
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   4572
   End
End
Attribute VB_Name = "ERKCVibor"
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
