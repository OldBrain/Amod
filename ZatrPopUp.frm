VERSION 5.00
Begin VB.Form ZatrPopUp 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   6336
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11508
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgTitleHelp 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      Height          =   192
      Left            =   0
      Picture         =   "ZatrPopUp.frx":0000
      ToolTipText     =   "О программе"
      Top             =   0
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Коментарий"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   228
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1176
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   5040
      Picture         =   "ZatrPopUp.frx":0542
      Top             =   1080
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   -1320
      Picture         =   "ZatrPopUp.frx":0C8C
      Top             =   360
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   360
      Left            =   3240
      Picture         =   "ZatrPopUp.frx":13D6
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   960
      Width           =   288
   End
End
Attribute VB_Name = "ZatrPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me, True
lblTitle.Caption = "Коментариии"
 
End Sub

