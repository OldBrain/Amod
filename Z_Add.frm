VERSION 5.00
Begin VB.Form Z_Add 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7350
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2880
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2880
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   720
      Width           =   4335
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      Caption         =   "Добавить"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Шифр затрат"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Категория расчета"
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
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   2415
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
      Height          =   240
      Left            =   0
      Picture         =   "Z_Add.frx":0000
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   960
      Picture         =   "Z_Add.frx":0542
      Top             =   0
      Width           =   285
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "123"
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   6810
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1320
      Picture         =   "Z_Add.frx":0C8C
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   480
      Picture         =   "Z_Add.frx":13D6
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Z_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Old As String
MakeWindow Me, True
lblTitle.Caption = "Добавление нового счета затрат"
Me.Label1.BackColor = RGB(207, 207, 207)
Me.Label2.BackColor = RGB(207, 207, 207)
'Me.Combo1.Text = z_sootn.

'Создаем комболист категорий расчета

Z_Sootn.rsCat.MoveFirst
Me.Combo1.Text = Z_Sootn.rsCat("Код") & "  " & Z_Sootn.rsCat("Name_Kategor")
Do While Not Z_Sootn.rsCat.EOF
'Me.Combo1.AddItem (CStr(Z_Sootn.rsCat("Код")) & vbTab & Z_Sootn.rsCat("Name_Kategor"))
Me.Combo1.AddItem (Z_Sootn.rsCat("Код") & "  " & Z_Sootn.rsCat("Name_Kategor"))
Z_Sootn.rsCat.MoveNext
Loop

'Создаем комболист счетов затрат

Z_Sootn.rsZat.MoveFirst
Me.Combo2.Text = Z_Sootn.rsZat("Schet") & "  " & Z_Sootn.rsZat("Schet_Name")
Do While Not Z_Sootn.rsZat.EOF
If Old <> Z_Sootn.rsZat("Schet_Name") Then Me.Combo2.AddItem (Z_Sootn.rsZat("Schet") & "  " & Z_Sootn.rsZat("Schet_Name"))
Old = Z_Sootn.rsZat("Schet_Name")
Z_Sootn.rsZat.MoveNext
Loop

End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub

Private Sub xpcmdbutton1_Click()
Dim name As String
Z_Sootn.rsZat.MoveFirst
Do While Not Z_Sootn.rsZat.EOF
If Z_Sootn.rsZat.Fields("Schet") = Val(Me.Combo2.Text) Then
If Z_Sootn.rsZat("Schet_Name") <> "" Then name = Z_Sootn.rsZat("Schet_Name")
End If

Z_Sootn.rsZat.MoveNext
Loop



Z_Sootn.rsZat.AddNew
Z_Sootn.rsZat.Fields("Kat") = Val(Me.Combo1.Text)
Z_Sootn.rsZat.Fields("Schet") = Val(Me.Combo2.Text)
Z_Sootn.rsZat.Fields("Schet_Name") = name
Z_Sootn.rsZat.UpdateBatch
'Z_Sootn.VS.Refresh
'Load Z_Sootn
Unload Me
End Sub
