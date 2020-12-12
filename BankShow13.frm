VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BankShow13 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9096
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   11868
   ControlBox      =   0   'False
   Icon            =   "BankShow13.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   758
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "BankShow13.frx":030A
      Left            =   7920
      List            =   "BankShow13.frx":030C
      TabIndex        =   7
      Text            =   "902"
      Top             =   480
      Width           =   3852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Разнести"
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   6
      Top             =   8640
      Width           =   1932
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11652
      _ExtentX        =   20553
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Нет соответствия л/сч"
            Key             =   "Key1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Все данные"
            Key             =   "Key2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Можно разнести"
            Key             =   "Key3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "XL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton Image1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Ок"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5772
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   11772
      _cx             =   20764
      _cy             =   10181
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"BankShow13.frx":030E
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
      DataMode        =   2
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   255
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Кад оплаты"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6840
      TabIndex        =   8
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   11055
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
      Picture         =   "BankShow13.frx":03F7
      Top             =   0
      Width           =   156
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
      Left            =   720
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10890
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   0
      Picture         =   "BankShow13.frx":0641
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BankShow13.frx":0D8B
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BankShow13.frx":14D5
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "BankShow13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim cRs As ADODB.Recordset
Dim RSD As ADODB.Recordset
Dim RsM As ADODB.Recordset
Dim rss As ADODB.Recordset
Dim estSc As Boolean
Dim S1 As Double
Dim Neo As String
Dim S2 As Double
Dim S3 As Double
Dim KZ As Integer

Private Sub BtnEnh11_Click()

End Sub

Private Sub Command1_Click()

' Добавляем данные в реестр документов
rs1.MoveFirst
Set rsDocReestr = New ADODB.Recordset
rsDocReestr.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.NachCod, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc"), Mconn, adOpenKeyset, adLockPessimistic
rsDocReestr.AddNew
rsDocReestr("Coment") = Me.lblTitle + " " + Combo1.Text
Cod = rsDocReestr("Cod") ' Код реестра
rsDocReestr("Data") = rs1("Дата платежа")
rsDocReestr("Nach") = Combo1.Text
rsDocReestr.UpdateBatch
rsDocReestr.Close


' Добавляем строки с оплатой в файл документов
Set RSD = New ADODB.Recordset
Set RsM = New ADODB.Recordset
Set rss = New ADODB.Recordset
rss.Open ("SELECT Settings.Neo FROM Settings"), Mconn
RSD.Open ("SELECT doc.Cod, doc.DataR, doc.KodN, doc.NameN, doc.KodKv, doc.NameKv, doc.Summa, doc.Key, doc.KeyAdding, doc.Stst, doc.Com, doc.Tip, doc.Button, doc.Dom, doc.RealData, doc.PLNOM FROM doc"), Mconn, adOpenKeyset, adLockPessimistic
RsM.Open ("SELECT MainOccupant.Numer, MainOccupant.kv_num, MainOccupant.Dom, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.OLDNUM FROM MainOccupant"), Mconn
'RsD.AddNew
Neo = rss("Neo")
rss.Close


Pod.Show
Pod.Label3.Caption = "Не удалось найти>"
Pod.Label3.Visible = False
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = KZ + 3
'MsgBox (Pod.ProgressBar1.Max)


' Перебераем строки файла импорта для определения номера л/сч
rs1.MoveFirst
Do While Not rs1.EOF
' Добавляем значения
If rs1("Ф И О") <> "" Then Pod.Label1.Caption = "Ищу >" + rs1("Ф И О")
Pod.Refresh


RSD.AddNew
RSD("DataR") = rs1("Дата платежа")
RSD("Cod") = Cod
RSD("NameKv") = rs1("Ф И О")
RSD("KodN") = Val(Me.Combo1.Text)

If MainForm.ErcFile = True Then RSD("Com") = Me.lblTitle + " за " + CStr(rs1("Период оплаты")) Else RSD("Com") = Me.lblTitle + " за " + rs1("Период оплаты")

RSD("NameN") = Me.Combo1.Text
RSD("Summa") = rs1("Сумма")
RSD("Stst") = 0
RSD("Tip") = "-"
'On Error GoTo ne
'RSD("RealData") = rs1("Период оплаты")
'ne:
'Теперь ищем соответствия номеров

RsM.MoveFirst
Do While Not RsM.EOF
If rs1("Счет") = "" Then rs1("Счет") = 0
If RsM("OLDNUM") = rs1("Счет") Then
'Если нашли

If rs1("Ф И О") <> "" Then Pod.Label2.Caption = "Добавляю оплату>" + rs1("Ф И О")
Pod.Refresh

RSD("KodKv") = RsM("Numer")
RSD("NameKv") = rs1("Ф И О") + "/" + RsM("FAM")
RSD("Dom") = RsM("DOM")
estSc = True
End If



RsM.MoveNext

Loop
'************************************
'Если нет то неопознанные суммы
If estSc = False Then
RSD("KodKv") = Neo
RSD("NameKv") = "Н/C" + "/" + rs1("Ф И О") + "/" + rs1("Адрес")
Pod.Label3.Visible = True
Pod.Label3.Caption = Pod.Label3.Caption + rs1("Ф И О")
RSD("Com") = Me.lblTitle + "/" + rs1("Ф И О") + "/" + rs1("Адрес") + " за " + rs1("Период оплаты")
RSD("Dom") = 1


Pod.Refresh
End If
estSc = False

Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
Pod.Refresh

rs1.MoveNext
Loop

RSD.UpdateBatch

Unload Pod

Unload Me
End Sub

Private Sub Command3_Click()
ВыводВExel
End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Me.Show
End Sub

Private Sub Form_Load()

Set rsDocReestr = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set cRs = New ADODB.Recordset

cRs.Open ("SELECT nachisleniy.Kod, nachisleniy.КодKategor, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy ORDER BY nachisleniy.Tip"), Mconn


' Заполняем combo для кода начисления
cRs.MoveFirst
'Set Combo1.DataSource = cRs
'Combo1.DataField = "Kod"

Do While Not cRs.EOF
Combo1.AddItem Str(cRs("Kod")) + "|" + cRs("Naim")
cRs.MoveNext
Loop




'ВСЕ данные

rs1.Open ("SELECT Bank.DATA AS [Дата платежа], Bank.LSCHET AS Счет, Bank.ADR AS Адрес, Bank.FIO AS [Ф И О], Bank.SUMMA AS Сумма, Bank.PERIODOPL AS [Период оплаты] From Bank ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic
KZ = 0
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Do While Not rs1.EOF
If MainForm.ErcFile = False Then S1 = S1 + rs1("Сумма")
rs1.MoveNext
KZ = KZ + 1
Loop
End If


' "Можно разнести"
rs2.Open ("SELECT Bank.DATA, Bank.LSCHET AS Счет, Bank.ADR AS Адрес, Bank.FIO AS [Ф И О], Bank.SUMMA AS Сумма, Bank.PERIODOPL AS [Период оплаты] From Bank WHERE (((Bank.LSCHET) Is Not Null Or (Bank.LSCHET)='0')) ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic

If rs2.RecordCount > 0 Then
rs2.MoveFirst
Do While Not rs2.EOF
S2 = S2 + rs2("Сумма")
rs2.MoveNext
Loop
End If

' Нет соответствия
rs3.Open ("SELECT Bank.DATA AS [Дата оплаты], Bank.SUMMA AS Сумма, Bank.FIO AS [Ф И О], Bank.ADR AS Адрес, Bank.LSCHET AS [Л/счет], Bank.PERIODOPL AS [Период оплаты] FROM Bank LEFT JOIN MainOccupant ON Bank.LSCHET = MainOccupant.OLDNUM WHERE (((MainOccupant.OLDNUM) Is Null))"), Mconn, adOpenKeyset, adLockPessimistic

If rs3.RecordCount > 0 Then
rs3.MoveFirst
Do While Not rs3.EOF
If MainForm.ErcFile = False Then S3 = S3 + rs3("Сумма")
rs3.MoveNext
Loop
End If

MakeWindow Me, True
FG1.Width = Me.Width / 15.40107
FG1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
Command1.Top = Image1.Top

'**************************************************************
'****************************************************************

'lblTitle = "Импорт оплаты из банка. Файл > " + BankImport.File1.FileName
'Label1.Caption = "Просмотр файла >" + BankImport.File1.FileName + ". Для продолжения нажмите <<Далее>>"

'Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S3)

lblTitle.Caption = lblTitle.Caption + "Общая суммареестра  > " + Str(BankShow.SummI)
'TabStrip1.Index = 1


Set FG1.DataSource = rs3
End Sub


Private Sub Image1_Click()

'Добавляем Realdata



'Msg.Show vbModal
'Msg.Label1.Caption = "Данные добавлены в реестр документов оплаты №" + Str(BankShow.Cod) + vbNewLine + "Сумма реестра банка=" + Str(BankShow.SummI) + vbNewLine + "Разнесенная сумма=" + Str(BankShow.SummI - S3) + vbNewLine + "Отклонение=" + Str(Round(S3, 2))
'Msg.Label1.Refresh

 'Unload ReestrDoc

'ReestrDoc.Fg.Refresh


Unload Me
Msg.Show
Msg.Label1.Caption = "Данные добавлены в реестр документов оплаты №" + Str(BankShow.Cod) + vbNewLine + "Сумма реестра банка=" + Str(BankShow.SummI) + vbNewLine + "Разнесенная сумма=" + Str(Round(BankShow.SummI - S3, 2)) + vbNewLine + "Отклонение=" + Str(Round(S3, 2))
Msg.Label1.Refresh


Unload BankShow

ReestrDoc.Show
ReestrDoc.Enabled = True
End Sub
Sub ВыводВExel()
   Const НачСтрока = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim i As Long, j As Long, K As Long, rДанные As String
   Dim v As Variant
   
   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
   
   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(FG1.Cols - 1) & FG1.Rows + НачСтрока
   ReDim v(FG1.Rows, FG1.Cols) 'Забыл указать
   
   If FG1.Rows > 0 Then
            For co = 1 To FG1.Cols - 1
         For rw = 0 To FG1.Rows - 1

             v(rw, co) = FG1.TextMatrix(rw, co)
             
         Next rw
         Next co
      ex1.Visible = True   'Еще забыл
      
      ws.Range(rДанные) = v
      
 End If
End Sub

Function XCol_(ByVal Column_ As Long) As String
    If (Column_ < 0) Then Column_ = 0
    If (Column_ < 26) Then
        XCol_ = Chr(Column_ + Asc("A"))
    ElseIf (Column_ < 676) Then
        XCol_ = Chr((Column_ \ 26) + Asc("A") - 1) & Chr((Column_ Mod 26) + Asc("A"))
    Else
        XCol_ = "ZZ"
    End If
End Function


Private Sub imgTitleHelp_Click()
Form2.Label1 = "   Импорт оплаты коммунальных услуг, из файла данных предоставленных банком. Это окно предназначено для предварительного просмотра импортируемых данных." + vbNewLine + "   Кроме того, Вы можете перенести данные в XL для дальнейшей печати"
Form2.Show
End Sub

Private Sub imgTitleMain_Click()
ChangeState Me
End Sub

Private Sub lblTitle_Click()
ChangeState Me
End Sub
Private Sub Form_Resize()
FG1.Width = Me.Width / 15.40107
   FG1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
'BtnEnh1.Top = Image1.Top

Command3.Left = Image1.Left + Image1.Width
End Sub


Private Sub TabStrip1_Click()



If TabStrip1.SelectedItem = "Все данные" Then
Set FG1.DataSource = rs1
Me.Command1.Enabled = False
Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S1)
End If

If TabStrip1.SelectedItem = "Можно разнести" Then
Me.Command1.Enabled = True
Set FG1.DataSource = rs2
Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S2)
End If

If TabStrip1.SelectedItem = "Нет соответствия л/сч" Then
Me.Command1.Enabled = False
Set FG1.DataSource = rs3
Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S3)
End If

'
'SelectedItem

End Sub

