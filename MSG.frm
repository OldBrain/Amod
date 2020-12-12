VERSION 5.00
Begin VB.Form MSG 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5184
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   6564
   ControlBox      =   0   'False
   Icon            =   "MSG.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2610
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
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
      Picture         =   "MSG.frx":038A
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   192
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
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
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   5850
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Сообщение"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   480
      Picture         =   "MSG.frx":08CC
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   720
      Picture         =   "MSG.frx":1016
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   240
      Picture         =   "MSG.frx":1760
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "MSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sts
Private Sub Command1_Click()
If sts <> "Реорг" Then
Unload Me
Else
sts = ""
Lgot_reorg.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
lblTitle = "Информационное окно"
MakeWindow Me, True
End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub
