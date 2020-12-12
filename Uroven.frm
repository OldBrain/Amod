VERSION 5.00
Begin VB.Form Uroven 
   ClientHeight    =   4896
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2952
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   4896
   ScaleWidth      =   2952
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option5 
      Caption         =   "1-5  колонка слева"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   2775
   End
   Begin VB.OptionButton Option4 
      Caption         =   "1-4  колонка слева"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ок"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "1-3  колонка слева"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "1-2  колонка слева"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1-я колонка слева"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Пожалуйста укажите количество уровней группировки Вашего отчета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Uroven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ur As Integer

Private Sub Command1_Click()
Pod.Show
Pod.Label1 = "Подождите, идет расчет!"
Pod.Refresh


Analizlgot.Об ur
'Anal_Zatrat.Об ur
End Sub

Private Sub Command2_Click()
Unload Me
Analizlgot.Enabled = True
End Sub

Private Sub Form_Load()
Option2.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Analizlgot.Show
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then ur = 1 Else ur = 10

End Sub
Private Sub Option2_Click()
If Option2.Value = True Then ur = 2 Else ur = 10
End Sub
Private Sub Option3_Click()
If Option3.Value = True Then ur = 3 Else ur = 10
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then ur = 4 Else ur = 10
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then ur = 5 Else ur = 10
End Sub

