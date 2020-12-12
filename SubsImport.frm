VERSION 5.00
Begin VB.Form SubsImport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4305
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Далее"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Отмена"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   2040
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
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "SubsImport.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   360
      Picture         =   "SubsImport.frx":074A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "SubsImport.frx":0E94
      Top             =   0
      Width           =   285
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   5
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1770
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
      Height          =   195
      Left            =   0
      Picture         =   "SubsImport.frx":15DE
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Укажите файл содержащий данные об оплате"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "SubsImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command2_Click()
SubsShow.Reestr = File1.FileName
SubsShow.Show
Unload Me
End Sub

Private Sub Dir1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + File1.FileName
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
'Dir1.Name = "*.dbf"
'Dir1.ListCount = "*.dbf"
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
File1.Pattern = "*.xls"
Label2.Caption = File1.Path + "\" + File1.FileName

End Sub

Private Sub Form_Unload(Cancel As Integer)
ReestrDoc.Enabled = True
End Sub

