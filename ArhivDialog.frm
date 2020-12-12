VERSION 5.00
Begin VB.Form ArhivDialog 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4665
   ClientLeft      =   2730
   ClientTop       =   3330
   ClientWidth     =   3675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   1215
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
      Picture         =   "ArhivDialog.frx":0000
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   3210
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "ArhivDialog.frx":024A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   480
      Picture         =   "ArhivDialog.frx":0994
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   240
      Picture         =   "ArhivDialog.frx":10DE
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "ArhivDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
ДисКоннект
MainForm.strDataName = "kvartplata.amd"
Коннект MainForm.strDataName
Unload Me
End Sub

Private Sub Form_Load()
lblTitle = "Выбор периода"
MakeWindow Me, True
File1.Path = App.Path + "\Data\Arhiv\"
File1.Path = Replace(File1.Path, "\\", "\")
File1.Pattern = "*.amd"

End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Show
End Sub

Private Sub OKButton_Click()
If File1.FileName <> "" Then

ДисКоннект
MainForm.strDataName = "\arhiv\" + File1.FileName
Коннект MainForm.strDataName
MainMenu.Command13.Caption = Left(File1.FileName, 8)
КоннектАрхив File1.FileName
Unload Me
Else
MsgBox "Вы не выбрали архив"
End If
End Sub
