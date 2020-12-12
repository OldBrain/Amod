VERSION 5.00
Begin VB.Form SubsImp 
   Caption         =   "Импорт субсидий"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form8"
   ScaleHeight     =   4560
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Адреса"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9495
   End
End
Attribute VB_Name = "SubsImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SubsAddr.Show

End Sub

