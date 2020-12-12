VERSION 5.00
Begin VB.Form Z_PathBest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6240
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   1  'CenterOwner
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "Отмена"
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
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "Ок"
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
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   5895
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
      Picture         =   "Z_pathBest.frx":0000
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   960
      Picture         =   "Z_pathBest.frx":0542
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
      Width           =   5730
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1320
      Picture         =   "Z_pathBest.frx":0C8C
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   480
      Picture         =   "Z_pathBest.frx":13D6
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Z_PathBest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
Dim P As String
Dim rsP As ADODB.Recordset
MakeWindow Me, True
lblTitle.Caption = "Укажите путь к Main.DBF"

Set rsP = New Recordset
rsP.Open ("Select path from z_nas"), Mconn
P = rsP("path")
'MsgBox (P)
If P = "" Then Drive1.Drive = "C:" Else Drive1.Drive = P
If P <> "" Then
Dir1.Path = P
'Label2.Caption = P
End If

'Dir1.Name = "*.dbf"
'Dir1.ListCount = "*.dbf"
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
File1.Pattern = "Main.dbf"
Label2.Caption = File1.Path + "\" + File1.FileName
MainForm.BestPath = File1.Path + "\" + File1.FileName
End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub
Private Sub Drive1_Change()
On Error GoTo DriveEr
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + File1.FileName
MainForm.BestPath = File1.Path + "\" + File1.FileName
DriveEr:
If Err.Number = 68 Then MsgBox "Нет диска в дисководе, или диск поврежден"
End Sub

Private Sub xpcmdbutton1_Click()
If File1.FileName = "" Then
MsgBox ("По указанному Вами пути файл Main.DBF не найден, пожалуйста уточните путь к файлу")
Else
Mconn.Execute ("UPDATE Z_Nas SET Z_Nas.Path = '" + Trim(Me.Label2.Caption) + "'")
MainForm.BestPath = Trim(Me.Label2.Caption)
Unload Me
End If
End Sub

Private Sub xpcmdbutton2_Click()
Unload Me
End Sub
