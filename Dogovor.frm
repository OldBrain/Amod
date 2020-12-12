VERSION 5.00
Begin VB.Form Dogovor 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3180
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Dogovor.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Строка коментария"
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "Окончание действия договора"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Начало действия договора"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Номер договора"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Пусто"
      DisabledPicture =   "Dogovor.frx":038A
      DownPicture     =   "Dogovor.frx":076F
      DragIcon        =   "Dogovor.frx":0B8B
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Абон.книжка"
      DownPicture     =   "Dogovor.frx":0F15
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Договор"
      DownPicture     =   "Dogovor.frx":0FBC
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Коментарий"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "по"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "От"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Договор №"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1560
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
      Height          =   240
      Left            =   0
      Picture         =   "Dogovor.frx":111D
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   240
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   2850
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   2160
      Picture         =   "Dogovor.frx":165F
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   480
      Picture         =   "Dogovor.frx":1DA9
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   120
      Picture         =   "Dogovor.frx":24F3
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Dogovor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDog As ADODB.Recordset
Dim rsSt As ADODB.Recordset
Dim Odin As Boolean

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
Odin = False
Me.KeyPreview = True
MakeWindow Me, True
lblTitle.Caption = "Абонкнижки / Договора"


Set rsDog = New ADODB.Recordset
rsDog.Open ("SELECT Dogovor.Numer, Dogovor.DocN, Dogovor.DogDataB, Dogovor.DogDataE, Dogovor.Com From Dogovor WHERE (((Dogovor.Numer)=" + Filter.Nm + "))"), Mconn, adOpenStatic, adLockPessimistic

Set rsSt = New ADODB.Recordset
rsSt.Open ("SELECT mainoccupant.Dog from mainoccupant WHERE (((mainoccupant.Numer)=" + Filter.Nm + "))"), Mconn, adOpenKeyset, adLockPessimistic



If Filter.FG.TextMatrix(Filter.FG.Row, 12) = 0 Then Option1.Value = True
If Filter.FG.TextMatrix(Filter.FG.Row, 12) = 1 Then Option2.Value = True
If Filter.FG.TextMatrix(Filter.FG.Row, 12) = 2 Then Option3.Value = True


If Option3.Value = False Then
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
End If



If rsDog.RecordCount <> 0 Then
If rsDog("DocN") <> "" Then Text1.Text = rsDog("DocN")
If rsDog("DogDataB") <> "" Then Text2.Text = rsDog("DogDataB")
If rsDog("DogDataE") <> "" Then Text3.Text = rsDog("DogDataE")
If rsDog("Com") <> "" Then Text4.Text = rsDog("Com")
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)


Filter.Enabled = True
If Option1.Value = True Then
rsSt("Dog") = 0
Filter.FG.TextMatrix(Filter.FG.Row, 12) = 0
Filter.FG.Cell(flexcpPicture, Filter.FG.Row, 12, Filter.FG.Row) = Filter.ImgNul
Filter.FG.Cell(flexcpBackColor, Filter.FG.Row, 12, Filter.FG.Row) = &H8080FF
If rsDog.RecordCount <> 0 Then
rsDog.Delete
rsDog.UpdateBatch
End If
End If

If Option3.Value = True Then
rsSt("Dog") = 2
Filter.FG.TextMatrix(Filter.FG.Row, 12) = 2
Filter.FG.Cell(flexcpPicture, Filter.FG.Row, 12, Filter.FG.Row) = Filter.ImgcellDog
Filter.FG.Cell(flexcpBackColor, Filter.FG.Row, 12, Filter.FG.Row) = &HE0E0E0


If rsDog.RecordCount = 0 Then
rsDog.AddNew
rsDog("Numer") = Filter.Nm
Text2.Text = rsDog("DogdataB")
Text2.Refresh
Text3.Text = rsDog("DogdataE")
Text3.Refresh
'rsDog.Requery

End If



'On Error GoTo er


rsDog("DocN") = Text1.Text
rsDog("DogdataB") = Text2.Text
rsDog("DogdataE") = Text3.Text
rsDog("Com") = Text4.Text
rsDog.UpdateBatch

Er:
If Err.Number <> 0 Then
MsgBox Err.Description
Err.Clear
End If

End If

If Option2.Value = True Then
rsSt("Dog") = 1
Filter.FG.TextMatrix(Filter.FG.Row, 12) = 1
Filter.FG.Cell(flexcpPicture, Filter.FG.Row, 12, Filter.FG.Row) = Filter.Imgcell
Filter.FG.Cell(flexcpBackColor, Filter.FG.Row, 12, Filter.FG.Row) = &HE0E0E0

If rsDog.RecordCount <> 0 Then
rsDog.Delete
rsDog.UpdateBatch
End If
End If






rsSt.UpdateBatch

'rsDog.Close


Filter.m_DS.m_RS.Close
Filter.m_DS.m_RS.Open "MainKV", Mconn, adOpenDynamic
'Filter.FG.Refresh

Filter.Enabled = True
End Sub

Private Sub imgTitleHelp_Click()
'Filter.Enabled = True

If Text1.Text = 0 And Option3.Value Then
If MsgBox("Договор без номера?", vbYesNo) = vbNo Then
'Dogovor.Show
Text1.SetFocus
Exit Sub
End If
End If


Unload Me
End Sub

Private Sub Option1_Click()
If rsDog.RecordCount <> 0 And Odin = False Then
Odin = True
If MsgBox("Данные договора (№,Дата...) бкдут удалены. Вы согласны?", vbYesNo) = vbNo Then
Option3.Value = True
End If
End If

If Option3.Value = False Then
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
End If

End Sub
Private Sub Option2_Click()

If rsDog.RecordCount <> 0 And Odin = False Then
Odin = True
If MsgBox("Данные договора (№,Дата...) бкдут удалены. Вы согласны?", vbYesNo) = vbNo Then
Option3.Value = True
End If
End If


If Option3.Value = False Then
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
End If

End Sub
Private Sub Option3_Click()


If Option3.Value Then

Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True

If rsDog.RecordCount = 0 Then
'rsDog.AddNew
Text2.Text = rsDog("DogdataB")
Text2.Refresh
Text3.Text = rsDog("DogdataE")
Text3.Refresh
End If


Else
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False

End If

End Sub


