VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form BImport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4308
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   7944
   ControlBox      =   0   'False
   Icon            =   "BImport.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4800
      Top             =   3840
      _ExtentX        =   804
      _ExtentY        =   804
      _Version        =   393216
      Protocol        =   2
      RemoteHost      =   "astlift.ru"
      RemotePort      =   21
      URL             =   "ftp://astlift_dbf:DgVjTURR@astlift.ru"
      Password        =   "DgVjTURR"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ОТПРАВИТЬ НА FTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   8
      Top             =   3840
      Width           =   2172
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Запись"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   1992
      Left            =   4200
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
      Left            =   240
      TabIndex        =   5
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   5970
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BImport.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BImport.frx":0B8C
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   0
      Picture         =   "BImport.frx":12D6
      Top             =   0
      Width           =   228
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
      Height          =   156
      Left            =   0
      Picture         =   "BImport.frx":1A20
      Top             =   0
      Width           =   156
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Укажите место на диске для записи файла"
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
Attribute VB_Name = "BImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command111_Click()

End Sub

Private Sub Command2_Click()

Pod.Show



Shell (App.Path + "\Util\rar.exe a " + Chr(34) + Label2.Caption + Replace(MenuNastr.StrNameB, "DBF", "RAR") + Chr(34) + " " + Chr(34) + App.Path + "\dbf\" + MenuNastr.StrNameB + Chr(34) + " -ep")

'MsgBox (App.Path + "\Util\rar.exe a " + Chr(34) + Label2.Caption + Replace(MenuNastr.StrNameB, "DBF", "RAR") + Chr(34) + " " + Chr(34) + App.Path + "\dbf\" + MenuNastr.StrNameB + Chr(34) + " -ep")

Label2.Caption = Replace(Label2.Caption, "\\", "\")
'  ChDir App.Path
'FileCopy "BASE_GH.RAR", Label2.Caption
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = 1001
Pod.Label1.Caption = "Архивирую, и записую в >" + Label2.Caption + Replace(MenuNastr.StrNameB, "DBF", "RAR")
Pod.Refresh
For i = 1 To 1000
Pod.ProgressBar1.Value = i

Next
MsgBox "Файл данных успешно записан в >" + Label2.Caption + Replace(MenuNastr.StrNameB, "DBF", "RAR")
Unload MenuNastr
Unload Pod
Unload Me
End Sub

Private Sub Command3_Click()
Dim FtpRec As ADODB.Recordset
Set FtpRec = New ADODB.Recordset
Set FtpRec.ActiveConnection = Mconn
FtpRec.Open ("SELECT Settings.NamePred, Settings.URL, Settings.UserName, Settings.Password FROM Settings")




Pod.Show
Pod.ProgressBar1.Max = 100000
Pod.ProgressBar1.Value = 10
rw = rw + 1



ArhName = Replace(Date, ".", "")
'+ "t" + Replace(Time, ":", "_")
Comment = "Дата создания_" + Replace(Date, ".", "") + "  Время создания_" + Replace(Time, ":", "_")
ftpName = Label2.Caption + ArhName + ".zip"

FtpPut = ftpName + " " + ArhName + ".zip"








'Shell (App.Path + "\Util\Pkzip.exe -a " + Chr(34) + Label2.Caption + ArhName + Chr(34) + " " + Chr(34) + App.Path + "\dbf\" + MenuNastr.StrNameB + Chr(34))

Shell (App.Path + "\Util\Pkzipc.exe -add " + Label2.Caption + ArhName + " " + App.Path + "\dbf\" + MenuNastr.StrNameB)

MsgBox ("Создаем архив -- " + App.Path + "\Util\Pkzipc.exe -add " + Label2.Caption + ArhName + " " + App.Path + "\dbf\" + MenuNastr.StrNameB)







'MsgBox ("Создаем архив" + FtpPut)

Pod.Show
For rw = 1 To 5000
Pod.ProgressBar1.Max = 5500
Pod.ProgressBar1.Value = rw
Pod.Label1 = "Создаем архив" + FtpPut
MainForm.WaiT (0.5)
Next

Pod.Hide

'If MsgBox("", vbYesNo) = vbNo Then Exit Sub



'For rw = 1 To 5000

'MainForm.WaiT (0.1)

'Next

ms = FtpRec("URL")


'Pkzip test.zip -ac license.doc





Pod.Show
Pod.ProgressBar1.Value = 1

With Inet1

 .URL = FtpRec("URL")
 '"ftp.astlift.ru"
 .UserName = FtpRec("UserName")
 '"astlift_dbf"
 .Password = FtpRec("Password")
  '"DgVjTURR"
 '.Execute , "POST", strDataFrom
 .Execute , "PUT " + FtpPut
 


End With

For rw = 1 To 5000
Pod.Label1 = "Отправляю файл " + FtpPut + " на сервер >" + FtpRec("URL")
MainForm.WaiT (0.3)
Pod.ProgressBar1.Value = rw
Next

Pod.Hide


FtpRec.Close

'For rw = 1 To 10000
'Next
'**********************************

'If FtpPutFile(RS&, "D:\index.htm", "index1.htm", 1, 0) = False Then MsgBox "Ошибка передачи файла!", vbExclamation


'Label2.Caption = Replace(Label2.Caption, "\\", "\")
'  ChDir App.Path
'FileCopy "BASE_GH.RAR", Label2.Caption
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = 1001
Pod.Label1.Caption = "Архивирую, и записую в >" + Label2.Caption + Replace(MenuNastr.StrNameB, "DBF", "RAR")
Pod.Refresh
'For i = 1 To 1000



'Pod.ProgressBar1.Value = i

'Next
MsgBox "Файл данных отправлен на " + ms
Unload MenuNastr
Unload Pod
Unload Me
End Sub

Private Sub Dir1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveEr
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")
DriveEr:
If Err.Number = 68 Then MsgBox "Нет диска в дисководе, или диск поврежден"
End Sub

Private Sub Drive1_LostFocus()
'Dir1.Path = Drive1.Drive
'File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
'Label2.Caption = File1.Path + "\" + File1.FileName

End Sub

Private Sub Form_Load()

 

lblTitle = "Запись файла с данными для банка"
MakeWindow Me, True
'Dat.Enabled = False
FileName = ""
'Dir1.Name = "*.dbf"
'Dir1.ListCount = "*.dbf"
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
'File1.Pattern = "*.dbf"
Label2.Caption = File1.Path + "\" + FileName

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dat.Enabled = True

End Sub

