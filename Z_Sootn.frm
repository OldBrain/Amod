VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Z_Sootn 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6120
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11475
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Z_Sootn.frx":0000
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   StartUpPosition =   1  'CenterOwner
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
      _extentx        =   2355
      _extenty        =   450
      caption         =   "Удалить"
      font            =   "Z_Sootn.frx":0275
   End
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
      _extentx        =   2355
      _extenty        =   450
      caption         =   "Добавить"
      font            =   "Z_Sootn.frx":02A1
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   11295
      _extentx        =   15055
      _extenty        =   661
      caption         =   "Закрыть"
      font            =   "Z_Sootn.frx":02CD
   End
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   11175
      _cx             =   19711
      _cy             =   7646
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Z_Sootn.frx":02F9
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   4
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Image Imgcell 
         Height          =   240
         Left            =   1320
         Picture         =   "Z_Sootn.frx":041E
         Top             =   3480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgNul 
         Height          =   480
         Left            =   480
         MouseIcon       =   "Z_Sootn.frx":0485
         Top             =   3840
         Width           =   480
      End
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
      Picture         =   "Z_Sootn.frx":078F
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   960
      Picture         =   "Z_Sootn.frx":0CD1
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
      Width           =   11130
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1320
      Picture         =   "Z_Sootn.frx":141B
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   480
      Picture         =   "Z_Sootn.frx":1B65
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Z_Sootn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBest As ADODB.Recordset
Public rsCat As ADODB.Recordset
Public rsZat As ADODB.Recordset
Public KeySh As Integer
Public Shet As String
Dim Cl As String

Private Sub Form_Load()

MakeWindow Me, True
lblTitle.Caption = "Расчет и анализ затрат на содержание ЖКХ"
' Открываем справочник статей затрат
Set rsCat = New ADODB.Recordset
Set rsZat = New ADODB.Recordset
 
'Rs_kat.CursorType = adOpenForwardOnly
'Rs_kat.LockType = adLockBatchOptimistic
 
 
rsZat.Open ("Select * from ShetBest"), Mconn, adOpenKeyset, adLockPessimistic

' Если файл пустой
'MsgBox rsZat.RecordCount

Mconn.Execute ("DELETE ShetBest.Schet_Name, ShetBest.* From ShetBest WHERE (((ShetBest.Schet_Name) Is Null))")

If rsZat.RecordCount <= 0 Then
Mconn.Execute ("INSERT INTO ShetBest ( Schet, Schet_Name, Kat, Best_Sh, Best_name, Ayn, Analit ) SELECT Schet.Schet, Schet.Schet_Name, 1, 0 AS Выражение1, '*' AS Выражение2, 0 AS Выражение3, 0 AS Выражение4 FROM Schet")
rsZat.Close
rsZat.Open ("Select * from ShetBest ORDER BY ShetBest.Schet_Name"), Mconn, adOpenKeyset, adLockPessimistic
End If
'*********************************





Set VS.DataSource = rsZat

VS.ComboSearch = flexCmbSearchAll

' Открываем main.dbf БЭСТ-4
Set rsBest = New ADODB.Recordset
Best
rsBest.Open ("Select Schet, Name_SCH, analit_Y_N from plan_sch Where status='1'"), BestConn

rsCat.Open ("Select код, Name_Kategor from Kategor"), Mconn


'Создаем комболист плана счетов
Cl = ""
rsBest.MoveFirst
Do While Not rsBest.EOF
Cl = Cl + CStr(rsBest("Schet")) & vbTab & rsBest("Name_SCH") + "|"
rsBest.MoveNext
Loop
VS.ColComboList(4) = Cl

'Создаем комболист категорий расчета
Cl = ""
rsCat.MoveFirst
Do While Not rsCat.EOF
Cl = Cl + CStr(rsCat("Код")) & vbTab & rsCat("Name_Kategor") + "|"
rsCat.MoveNext
Loop
VS.ColComboList(3) = Cl

'*****************КАРТИНКИ ДЛЯ АНАЛИТИКИ********************

For R = 1 To VS.Rows - 1
      If VS.TextMatrix(R, 6) = 0 Then
  VS.Cell(flexcpPicture, R, 0, R) = ImgNul
   VS.ComboList = ""
   VS.Cell(flexcpBackColor, R, 7, R) = vbWhite
   End If
   
   If VS.TextMatrix(R, 6) = -1 Then
   
  VS.ComboList = "..."
  VS.Cell(flexcpPicture, R, 0, R) = Imgcell
  VS.Cell(flexcpBackColor, R, 7, R) = &HE0E0E0
  
   End If
                     Next

'*************************************
End Sub

Private Sub imgTitleHelp_Click()
xpcmdbutton1_Click
End Sub

Private Sub VS_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' Заполняем колонку 5 /Наименование счета/
If Col = 4 Then
rsBest.MoveFirst
Do While Not rsBest.EOF
If VS.TextMatrix(Row, Col) = rsBest("Schet") Then
VS.TextMatrix(Row, 5) = rsBest("Name_SCH")
'If rsBest("analit_Y_N") = ttue Then
'MsgBox rsBest("analit_Y_N")
VS.TextMatrix(Row, 6) = rsBest("analit_Y_N")
End If
rsBest.MoveNext
Loop
End If


' КАРТИНКИ ДЛЯ АНАЛИТИКИ

For R = 1 To VS.Rows - 1
      If VS.TextMatrix(R, 6) = 0 Then
  VS.Cell(flexcpPicture, R, 0, R) = ImgNul
   VS.ComboList = ""
   VS.Cell(flexcpBackColor, R, 7, R) = vbWhite
   End If
   
   If VS.TextMatrix(R, 6) = -1 Then
   
  VS.ComboList = "..."
  VS.Cell(flexcpPicture, R, 0, R) = Imgcell
  VS.Cell(flexcpBackColor, R, 7, R) = &HE0E0E0
  
   End If
                     Next

End Sub
Private Sub VS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)



Z_Anal.Show (1)
End Sub



Private Sub VS_Click()
'Ключ для связи с файлом аналитики
KeySh = VS.TextMatrix(VS.Row, 8)
Shet = VS.TextMatrix(VS.Row, 4)

If VS.TextMatrix(VS.Row, 6) = True And VS.Col = 7 Then
'MsgBox ("")
VS.ComboList = "..."
Else
'Cancel = True
VS.ComboList = ""
End If
End Sub

Private Sub VS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
' Запрет редактирования колонок
If Col = 5 Or Col = 1 Then Cancel = True

If VS.TextMatrix(Row, 6) = True And Col = 7 Then
'MsgBox ("")
VS.ComboList = "..."
Else
'Cancel = True
VS.ComboList = ""
End If



If Col = 7 And VS.TextMatrix(Row, 6) = False Then Cancel = True
End Sub


Private Sub xpcmdbutton1_Click()
Unload Me
End Sub

Private Sub xpcmdbutton2_Click()
Z_Add.Show 1
End Sub
