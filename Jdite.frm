VERSION 5.00
Begin VB.Form Jdite 
   ClientHeight    =   1128
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8616
   ControlBox      =   0   'False
   Icon            =   "Jdite.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MouseIcon       =   "Jdite.frx":030A
   ScaleHeight     =   1128
   ScaleWidth      =   8616
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Пожалуйста подождите"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8628
   End
End
Attribute VB_Name = "Jdite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.KeyPreview = True
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
      Unload Me
   End If
End Sub

