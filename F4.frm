VERSION 5.00
Begin VB.Form F4 
   Caption         =   "Тариф"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form4"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
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
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   2772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Тариф"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2772
   End
End
Attribute VB_Name = "F4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ZPlan.T = Me.Text1.Text
ZPlan.Command1.Caption = "Тариф =" + ZPlan.T + " Изменить?"
Unload Me
End Sub

