VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Pod 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2784
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4704
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawWidth       =   100
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
      MousePointer    =   5
      Min             =   1
      Max             =   1,00000e5
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   1560
      Picture         =   "Pod.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000002&
      BorderWidth     =   6
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   230
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      X1              =   390
      X2              =   390
      Y1              =   0
      Y2              =   230
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000002&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   6
      X1              =   391
      X2              =   1
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1272
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4380
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
      Height          =   168
      Left            =   4200
      Picture         =   "Pod.frx":0400
      Top             =   120
      Width           =   168
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ожидайте окончания операции"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   6
      X1              =   390
      X2              =   0
      Y1              =   230
      Y2              =   230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      BackStyle       =   0  'Transparent
      Caption         =   "Пожалуйста подождите"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4332
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Pod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Doc.Enabled = True
Unload Me
End Sub

Private Sub Form_Activate()
DoEvents
    
    'тут типа SQL
    'PauseTime = 2
    'Start = Timer
    'Do While Timer < Start + PauseTime
   
     '   ProgressBar1.Value = i
      '  i = i + 10
        
     '   DoEvents
   'Loop
    'ну и хватит пожалуй юзера мучать
    'Unload Me
End Sub

Private Sub Form_Load()
'ProgressBar1.Value = 50
DoEvents
End Sub

