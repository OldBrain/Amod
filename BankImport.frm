VERSION 5.00
Begin VB.Form BankImport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4308
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   7944
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "BankImport.frx":0000
   NegotiateMenus  =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   1920
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Переименовать файлы *.625 в *.DBF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   9
      Top             =   1080
      Width           =   3492
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ОТПРАВИТЬ FTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3120
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Далее >>"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Отмена"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1800
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3615
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
      Height          =   180
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "BankImport.frx":030A
      Top             =   0
      Width           =   180
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Окно выбора фыйлов оплаты банка"
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
      TabIndex        =   7
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   7860
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BankImport.frx":0718
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BankImport.frx":0E62
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   720
      Picture         =   "BankImport.frx":15AC
      Top             =   0
      Width           =   228
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Укажите файл содержащий данные об оплате, и нажмите ""Далее>>"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "BankImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public var As String

Private Sub Command1_Click()
MainMenu.Enabled = True
Unload Me
End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command2_Click()
Command2.Enabled = False
' Укарачиваем имя файла до 6 символов
'MsgBox (File1.Path + Right(File1.FileName, 10))
FileCopy File1.Path + "\" + File1.FileName, File1.Path + "\tmp" + Right(File1.FileName, 7)
'App.Path "/Dbf/BETEM.DBF", App.Path + "/dbf/" + "BET.DBF"

If var = "TSG" Then
'SCH_ET.Show
Unload Me
Exit Sub
End If

Me.Label1.ForeColor = vbRed

Me.Label1.Caption = "Пожалуйста подождите"
Me.Label1.Refresh

Me.Enabled = False

BankShow.Reestr = File1.FileName
BankShow.Show 1
'BankShow.Reestr = BankImport.File1.FileName

Unload Me
End Sub

Private Sub Command21_Click()

End Sub

Private Sub Command3_Click()
'BankPOLE.DBFName = File1.FileName
If File1.FileName <> "" Then BankPOLE.DBFName = File1.Path + "\" + File1.FileName
'BankPOLE.Show
'BankShow.Reestr = BankImport.File1.FileName

Unload Me
End Sub

Private Sub Command4_Click()
'Переименовуем файлы
'MsgBox ("cmd ren " + File1.Path + "\" + "*.625 *.dbf")

Shell "cmd /c ren " + File1.Path + "\" + "*.625 *.dbf"
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Me.File1.Refresh

End Sub

Private Sub Dir1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + File1.FileName



'oWsh.Run 'REN *.625 *.DBF'
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveEr
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + File1.FileName
DriveEr:
If Err.Number = 68 Then MsgBox "Нет диска в дисководе, или диск поврежден"
End Sub

Private Sub Drive1_LostFocus()
'Dir1.Path = Drive1.Drive
'File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
Label2.Caption = File1.Path + "\" + File1.FileName

End Sub

Private Sub Form_Load()
lblTitle = "Импорт оплаты из файлов предоставленных банком"
MakeWindow Me, True
ReestrDoc.Enabled = False
Drive1.Drive = "C:"

'Dir1.Name = "*.dbf"
'Dir1.ListCount = "*.dbf"
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
File1.Pattern = "*.dbf"
Label2.Caption = File1.Path + "\" + File1.FileName


End Sub

Private Sub Form_Unload(Cancel As Integer)
'If var <> "TSG" Then
ReestrDoc.Enabled = True
End Sub

Private Sub imgTitleHelp_Click()
Command1_Click
End Sub

