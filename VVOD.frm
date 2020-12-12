VERSION 5.00
Begin VB.Form VVOD 
   Caption         =   "Новая запись"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form7"
   ScaleHeight     =   1260
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ввод"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "VVOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 Schet.Hide
 If VVOD.Text1.Text <> Null Then
 Schet.Rs_kat.AddNew
 Schet.Rs_kat.Fields("Schet") = VVOD.Text1.Text
Schet.Rs_kat.Fields("Schet_Name") = VVOD.Text2.Text
Schet.Rs_kat.UpdateBatch
End If
Unload Me
Schet.Show
Schet.Enabled = True
Schet.FG.DataRefresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Schet.FG.DataRefresh
Schet.Show
Schet.Enabled = True
End Sub
