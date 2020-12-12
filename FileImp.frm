VERSION 5.00
Begin VB.Form FileImp 
   Caption         =   "Импорт льгот"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form8"
   ScaleHeight     =   6525
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   6135
      Left            =   4080
      Pattern         =   "*.DBF"
      TabIndex        =   5
      Top             =   240
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ок"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Укажите путь к файлам с данными о льготах"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "FileImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public ImpPath As String
Dim ImpF1, ImpF2 As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
i = 0
Do While File1.List(i) <> ""
'MsgBox File1.List(i)
If InStr(1, File1.List(i), "GEK") <> 0 Then ImpF1 = File1.List(i)
If InStr(1, File1.List(i), "BAZA") <> 0 Then ImpF2 = File1.List(i)
i = i + 1
'MsgBox ImpF1
Loop
If ImpF1 = "" Then
MsgBox ("Нет файла GEK*.DBF")
Exit Sub
End If
If ImpF2 = "" Then
MsgBox ("Нет файла BAZA*.DBF")
Exit Sub
End If

ImpPath = File1.Path
FileCopy ImpPath + "\" + ImpF1, App.Path + "\Import\GEK.dbf"
FileCopy ImpPath + "\" + ImpF2, App.Path + "\Import\baza.dbf"

MsgBox "Первый шак завершен, файлы успешно скопированы в директорию " + App.Path + "\Import"
ImpLg.Command2.Enabled = True
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Dir1.Refresh
End Sub

