VERSION 5.00
Begin VB.Form SposobR1 
   Caption         =   "Способ расчета"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Без учета исправлений"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "С учетом исправлений"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "SposobR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Label1 = "Пожалуйста подождите. Идет расчет."
Lic.SPR = "СУчетом"
'С учетом исправлений
For Rw = 1 To Lic.FG1.Rows - 1
Lic.НовыйРасчет1 Rw, Lic.FG1.TextMatrix(Rw, 26), Lic.SPR
Next
Unload Me

End Sub

Private Sub Command2_Click()
Label1 = "Пожалуйста подождите. Идет расчет."
'Без учета исправлений
Lic.SPR = "БезУчета"
For Rw = 1 To Lic.FG1.Rows - 1
Lic.НовыйРасчет1 Rw, Lic.FG1.TextMatrix(Rw, 26), Lic.SPR
'Lic.FG1.TextMatrix(Rw, 27) = 0
Next

Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()

Lic.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Lic.Enabled = True
Lic.FG1.DataRefresh
End Sub

