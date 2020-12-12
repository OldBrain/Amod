VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form BankShow 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9588
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   11868
   ControlBox      =   0   'False
   Icon            =   "BankShow.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   799
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   10320
      TabIndex        =   9
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton Image1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Отмена"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   1455
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00808080&
      Caption         =   "Далее>>"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00404040&
      Caption         =   "Далее>>>"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Далее >>>>"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton BtnEnh4 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Далее>>>"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   9495
      _cx             =   16748
      _cy             =   3625
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
      FormatString    =   $"BankShow.frx":0442
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
      DataMode        =   3
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
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   720
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
      Picture         =   "BankShow.frx":0524
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
      TabIndex        =   1
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10890
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   0
      Picture         =   "BankShow.frx":076E
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BankShow.frx":0EB8
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BankShow.frx":1602
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "BankShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBFConn1 As ADODB.Connection
'Dim mconn As ADODB.Connection
Public dbfRs As ADODB.Recordset
Dim AccessRs As ADODB.Recordset
Dim ErrorStst As Integer
Dim rsDoobl As ADODB.Recordset
Dim RsSet As ADODB.Recordset
Dim rsNul As ADODB.Recordset
Dim rsOk As ADODB.Recordset
Dim Shag As Integer
Dim rsDoc As ADODB.Recordset
Dim rsDocReestr As ADODB.Recordset
Public Cod As Integer
Dim rsNas As ADODB.Recordset
Dim rsReestr As ADODB.Recordset
Dim Neo As String
Public Reestr As String
Dim DItem As String
Dim Bn As String
Public For_Dell As String ' имя временного файла для удаления
Dim Clik As Integer
Public s As Double
Dim Beg As Boolean
Dim Ssum As Double, Kv As Double, Musor As Double, Lift As Double, Otopl As Double, Gv As Double, Hv As Double, El As Double, Sliv As Double
Dim rsEnd As ADODB.Recordset
Public SummI As Double
'Dim FN As String
Dim KORP As String
'Dim Old As Boolean
Public NewN As Boolean
Public ops As Boolean
Public TSG As Boolean

Dim rsBank As ADODB.Recordset
Dim rsDbfBank As ADODB.Recordset




Private Sub BtnEnh1_1_Click()

End Sub

Private Sub BtnEnh1_Click()
DoEvents

'MsgBox NewN

'MsgBox Old



Shag = 1

If ErrorStst <> 0 Then
Msg.Show
Msg.Label1.Caption = "Кол-во ошибок =" + Str(ErrorStst) + vbNewLine + " Сначала исправте ошибки!" + vbNewLine + "Ячейки с ошибками выделены цветом:" + vbNewLine + "ЖЁЛТЫЙ - это несоответствие общей сумы платежа, сумме платежей по каждой категории оплаты, необходимо проставить правильные суммы оплаты по категориям" + vbNewLine + "КРАСНЫЙ - это пустые значения полей. Пока Вы не заполните пустые значения данными продолжение операции импорта будет невозможно. Если Вы не знаете что проставить, можно ввести любые значения, но не забывать об этом в дальнейшем."
Shag = 0
Exit Sub
End If


'Set Mconn = New ADODB.Connection
 ' Mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' Mconn.Open "data/Kvartplata.mdb"
'mconn.Execute "SELECT '' AS DATA, '' AS KFOSB, '' AS KISP, '' AS NOP, '' AS SUMMA, '' AS SBOR, '' AS FIO, '' AS ADR, '' AS LSCHET, '' AS PERIODOPL, '' AS SKOMM, '' AS SLIFT, '' AS SMUSOR, '' AS SELEN, '' AS SGVS, '' AS STEPLO, '' AS SHVODA, '' AS SSLIV, '' AS PLNOM, '' AS PLDATE, '' AS NRS INTO Bank"
'mconn.Execute "SELECT BankShablon.Key, BankShablon.DATA, BankShablon.KFOSB, BankShablon.SUMMA, BankShablon.SBOR, BankShablon.FIO, BankShablon.ADR, BankShablon.LSCHET, BankShablon.PERIODOPL, BankShablon.KV, BankShablon.LIFT, BankShablon.MUSOR, BankShablon.SELEN, BankShablon.GVoda, BankShablon.Otopl, BankShablon.HVoda, BankShablon.SSLIV, BankShablon.PLNOM, BankShablon.PLDATE, BankShablon.NRS INTO Bank FROM BankShablon"

Mconn.Execute "DELETE Bank.* FROM Bank"

AccessRs.Open ("SELECT Bank.DATA, Bank.KFOSB, Bank.SUMMA, Bank.SBOR, Bank.FIO, Bank.ADR, Bank.LSCHET, Bank.PERIODOPL, Bank.KV,Bank.LIFT, Bank.MUSOR, Bank.SELEN, Bank.GVoda, Bank.Otopl, Bank.HVoda, Bank.SSLIV, Bank.PLNOM, Bank.PLDATE, Bank.NRS, Bank.NewNum FROM Bank"), Mconn, adOpenDynamic, adLockBatchOptimistic



'If mewn = True Then AccessRs.Open ("SELECT Bank.DATA, Bank.KFOSB, Bank.SUMMA, Bank.SBOR, Bank.FIO, Bank.ADR, Bank.LSCHET, Bank.PERIODOPL, Bank.KV,Bank.LIFT, Bank.MUSOR, Bank.SELEN, Bank.GVoda, Bank.Otopl, Bank.HVoda, Bank.SSLIV, Bank.PLNOM, Bank.PLDATE, Bank.NRS, Bank.NewNum FROM Bank"), Mconn, adOpenDynamic, adLockBatchOptimistic

'dbfRs.UpdateBatch

'Дополняем нолями номера лиц счетов

dbfRs.MoveFirst
              Do While Not dbfRs.EOF
Bn1 = dbfRs("NewNum")
Do While Len(Bn1) < 12
Bn1 = "0" + Bn1
Loop
dbfRs("NewNum") = Bn1
dbfRs.Update
                   dbfRs.MoveNext
                   Loop
'*****************



dbfRs.MoveFirst
Do While Not dbfRs.EOF
AccessRs.AddNew


' Если это файл из ЕРКЦ
If MainForm.ErcFile = True Then
AccessRs("Data") = dbfRs("opldate") ' Дата платежа
AccessRs("Summa") = dbfRs("N_Sum") ' Общая сумма оплаты
AccessRs("FIO") = dbfRs("FAMILY") 'Фамилия имя отчество
If IsNull(dbfRs("KORP")) Then
KORP = ""
Else
KORP = Str(dbfRs("KORP"))
End If
AccessRs("Adr") = dbfRs("Str") + " д." + Str(dbfRs("House_Num")) + "корп." + KORP + " кв." + Str(dbfRs("Room_Num")) 'Адрес
'AccessRs("LSCHET") = dbfRs("CHECKNUM") ' Лиц.счет OLDNUM
AccessRs("LSCHET") = dbfRs("AB_ID") ' Лиц.счет OLDNUM
AccessRs("PLNOM") = dbfRs("NPP") 'НОМЕР ПЛАТЕЖКИ
AccessRs("PERIODOPL") = dbfRs("AB_DO") 'ПЕРИОД ОПЛАТЫ
End If



If MainForm.Old = True And MainForm.ErcFile = False Then

'If dbfRs("Date") <> "" Then AccessRs("Data") = dbfRs("Date") Else AccessRs("Data") = Date
If dbfRs("Date") <> "" Then AccessRs("Data") = dbfRs("Date") Else AccessRs("Data") = Date
Else
'If MainForm.ErcFile = False Then If dbfRs("Data") <> "" Or dbfRs("Data") = "0" Then AccessRs("Data") = dbfRs("Data") Else AccessRs("Data") = Date

If MainForm.ErcFile = False Then If dbfRs("Data") <> "" Or dbfRs("Data") = "0" Then AccessRs("Data") = Date Else AccessRs("Data") = Date


End If

If MainForm.ErcFile = False Then AccessRs("KFOSB") = dbfRs("KFOSB")

If MainForm.Old = True Then
If MainForm.ErcFile = False Then AccessRs("Summa") = dbfRs("Sopl") ' Общая сумма оплаты
Else
AccessRs("Summa") = dbfRs("Summa") ' Общая сумма оплаты
End If
'AccessRs("Sbor") = dbfRs("Sbor") ' Скорее всего ком.сбор банка
If MainForm.ErcFile = False Then AccessRs("FIO") = dbfRs("FIO") 'Фамилия имя отчество
If MainForm.ErcFile = False Then AccessRs("Adr") = dbfRs("Adr") 'Адрес





If MainForm.ErcFile = False Then AccessRs("LSCHET") = dbfRs("LSCHET") ' Лиц.счет OLDNUM


If MainForm.ErcFile = False Then If dbfRs("PERIODOPL") <> "" Then AccessRs("PERIODOPL") = dbfRs("PERIODOPL") 'Период за который платит квартиросъемщик

If MainForm.Old = False Then
If dbfRs("SKOMM") <> "" Then AccessRs("KV") = Replace(dbfRs("SKOMM"), ",", ".") 'КВАРТПЛАТА
If dbfRs("SLift") <> "" Then AccessRs("Lift") = Replace(dbfRs("SLift"), ",", ".") 'ЛИФТ
If dbfRs("SMUSOR") <> "" Then AccessRs("MUSOR") = Replace(dbfRs("SMUSOR"), ",", ".") 'МУСОР
If dbfRs("Selen") <> "" Then AccessRs("selen") = Replace(dbfRs("selen"), ",", ".") '
If dbfRs("SGVS") <> "" Then AccessRs("Gvoda") = Replace(dbfRs("SGVS"), ",", ".") 'ГОРЯЧАЯ ВОДА
If dbfRs("Steplo") <> "" Then AccessRs("Otopl") = Replace(dbfRs("STEPLO"), ",", ".") 'ОТОПЛЕНИЕ
If dbfRs("SHVODA") <> "" Then AccessRs("HVoda") = Replace(dbfRs("SHVODA"), ",", ".") 'ХОЛОДНАЯ ВОДА
If dbfRs("SSLIV") <> "" Then AccessRs("SSLIV") = Replace(dbfRs("SSLIV"), ",", ".") 'СЛИВ
AccessRs("PLNOM") = Replace(dbfRs("PLNOM"), ",", ".") 'НОМЕР ПЛАТЕЖА
AccessRs("PLDATE") = Replace(dbfRs("PLDATE"), ",", ".") 'ДАТА ПЛАТЕЖА

AccessRs("NRS") = Replace(dbfRs("NRS"), ",", ".") ' Расчетный счет

Else

If MainForm.ErcFile = False Then If dbfRs("Sopl") <> "" Then AccessRs("Lift") = Replace(dbfRs("Sopl"), ",", ".") 'КВАРТПЛАТА

End If
If NewN = True Then AccessRs("NewNum") = dbfRs("Newnum")


If NewN = True Then fg1.Cols = 14
AccessRs.UpdateBatch
dbfRs.MoveNext
Loop



'*****/////****
AccessRs.Requery
AccessRs.MoveFirst
Do While Not AccessRs.EOF
If Len(AccessRs("LSCHET")) = 12 Then
AccessRs("NewNum") = AccessRs("LSCHET")

' Пробуем убрать обноление
'AccessRs("LSCHET") = "0"

AccessRs.UpdateBatch
End If
AccessRs.MoveNext
Loop
AccessRs.Requery
'******//////*****
fg1.ColHidden(2) = False
fg1.ColHidden(3) = False
fg1.ColHidden(4) = False


' Если это файлы для лифтов Банк или ЕРЦ сразу переходим на ворму 13, текущее окно закрываем, разносить будем в форме 13



If MainForm.LiftFile = True Or MainForm.ErcFile = True Then

BankShow13.Show

If MainForm.LiftFile = True Then BankShow13.lblTitle = "Файл банка " + Fn
If MainForm.ErcFile = True Then BankShow13.lblTitle = "Файл ЕРКЦ " + Fn

Unload BankShow
End If

If NewN = True Then


rsDocReestr.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.NachCod, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc"), Mconn, adOpenKeyset, adLockPessimistic


rsDocReestr.AddNew
rsDocReestr("Coment") = "Реестр банка " + Reestr + " новые номера л/сч."
Cod = rsDocReestr("Cod")
rsDocReestr("Data") = AccessRs("Data")
rsDocReestr("Nach") = "Любое начисление"
rsDocReestr.UpdateBatch
rsDocReestr.Close

If MainForm.Bank12 = 0 Then BankShowNew.Show vbModal


If MainForm.Bank12 = 1 Then
Me.Enabled = False
Me.Visible = False

BankShowNew.Show
End If

End If






DBFConn1.Close

Set rsDoobl = New ADODB.Recordset
rsDoobl.Open ("SELECT BankSvyz.OLDNUM, BankSvyz.FAM, BankSvyz.IM, BankSvyz.OT, BankSvyz.КОД, BankSvyz.NAIM_KLS, BankSvyz.Num, BankSvyz.LSCHET, BankSvyz.FIO, BankSvyz.ADR, BankSvyz.KV, BankSvyz.LIFT, BankSvyz.MUSOR, BankSvyz.SELEN, BankSvyz.GVoda, BankSvyz.Otopl, BankSvyz.HVoda, BankSvyz.SSLIV, BankSvyz.Key, BankSvyz.DATA, BankSvyz.numer, BankSvyz.PERIODOPL From BankSvyz WHERE (((BankSvyz.OLDNUM) In (SELECT [OLDNUM] FROM [BankSvyz] As Tmp GROUP BY [OLDNUM],[KV],[LIFT],[MUSOR],[SELEN],[GVoda],[Otopl],[HVoda],[SSLIV] HAVING Count(*)>1  And [KV] = [BankSvyz].[KV] And [LIFT] = [BankSvyz].[LIFT] And [MUSOR] = [BankSvyz].[MUSOR] And [SELEN] = [BankSvyz].[SELEN] And [GVoda] = [BankSvyz].[GVoda] And [Otopl] = [BankSvyz].[Otopl] And [HVoda] = [BankSvyz].[HVoda] And [SSLIV] = [BankSvyz].[SSLIV]))) ORDER BY BankSvyz.OLDNUM, BankSvyz.KV, BankSvyz.LIFT, BankSvyz.MUSOR, BankSvyz.SELEN, BankSvyz.GVoda, BankSvyz.Otopl, BankSvyz.HVoda, BankSvyz.SSLIV"), Mconn




Set fg1.DataSource = rsDoobl
'rsDoobl.Close

If Shag = 1 Or Shag = 2 Then
'FG1.Cell(flexcpBackColor, 1, 1, FG1.Rows - 1, 8) = &H80000018
'FG1.Cell(flexcpBackColor, 1, 8, FG1.Rows - 1, FG1.Cols - 1) = RGB(200, 255, 200)
End If
'Назначаем нулевую колонку для крыжей



For rw = 1 To fg1.Rows - 1
fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked
'flexChecked
Next rw

BtnEnh1.Visible = False
BtnEnh2.Visible = True
Clik = 0


For rw = 1 To fg1.Rows - 1
If Compare(UCase(fg1.TextMatrix(rw, 9)), UCase(fg1.TextMatrix(rw, 2)), 5) > 0.7 Then
'FG1.Cell(flexcpBackColor, Rw, 1, Rw, FG1.Cols - 1) = RGB(200, 255, 100)
fg1.Cell(flexcpChecked, rw, 0) = flexChecked
Else
'FG1.Cell(flexcpBackColor, Rw, 1, Rw, FG1.Cols - 1) = vbRed
fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked
End If
Next rw


'FG1.Refresh

For rw = 1 To fg1.Rows - 1
Nitem = fg1.TextMatrix(rw, 1)
If rw = fg1.Rows - 1 Then Exit For
If Nitem = fg1.TextMatrix(rw + 1, 1) And fg1.Cell(flexcpChecked, rw + 1, 0) = flexUnchecked And fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked Then
fg1.Cell(flexcpBackColor, rw, 1, rw + 1, 10) = RGB(100, 150, 100)
'FG1.Cell(flexcpBackColor, Rw + 1, 1, Rw + 1, FG1.Cols - 1) = RGB(100, 150, 100)
fg1.Cell(flexcpFontBold, rw, 1, rw + 1, fg1.Cols - 1) = True
'FG1.Cell(flexcpFontBold, Rw, 1, Rw + 1, FG1.Cols - 1) = True
'FG1.Refresh

End If
Next rw

'Me.Refresh
For rw = 1 To fg1.Rows - 1

Next rw

If MainForm.Bank12 <> 1 Then
Msg.Show
Msg.Label1.Caption = "Ошибки в суммах исправлены" + vbNewLine + "Теперь Вам будет предложен список фамилий, которые программе не удалось идентифицировать однозначно. Т.е. в реестре банка есть номера лицевых счетов, которые не являются уникальными для Вашей базы. Вам необходимо отметить квартиросъемщиков в лицевой счет которым должна попасть данная сумма оплаты."
Msg.Label1.Caption = Msg.Label1.Caption + vbNewLine + "!ВНИМАНИЕ ВАЖНО!" + vbNewLine + " Обратите внимание, что в списке присутствует ДВЕ ИЛИ БОЛЕЕ строки соответствующие ТОЛЬКО ОДНОЙ оплате. Ваша задача выбрать только ОДНУ единственно верную."
End If

End Sub

Private Sub BtnEnh2_Click()




Clik = Clik + 1
If Clik <= 1 Then
rsDocReestr.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.NachCod, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc"), Mconn, adOpenKeyset, adLockPessimistic
rsDocReestr.AddNew

AccessRs.Requery

rsDocReestr("Coment") = "Реестр банка " + Reestr
Cod = rsDocReestr("Cod")
rsDocReestr("Data") = AccessRs("Data")
rsDocReestr("Nach") = "Любое начисление"
rsDocReestr.UpdateBatch
rsDocReestr.Close
'BtnEnh2.Visible = False

End If

'Clik = Clik + 1
If fg1.Rows <= 1 Then
Shag = 2
Msg.Show
Msg.Label1.Caption = "Следующее окно, это список строк реестра банка, для которых программе неудалось найти соответствующих номеров лицевх счетов среди Ваших квартиросъемщиков."
Msg.Label1.Caption = Msg.Label1.Caption + vbNewLine + "Возможно в реестре банка допущена ошибка в номере лицевого счета, или этот номер лиц.счета не относится к Вашему ЖЭК."


Shag = 2
'FG1.AutoResize = True

Set rsNul = New ADODB.Recordset
rsNul.Open ("SELECT BankSvyz.numer, BankSvyz.FAM, BankSvyz.IM, BankSvyz.OT, BankSvyz.КОД, BankSvyz.NAIM_KLS, BankSvyz.Num, BankSvyz.LSCHET, BankSvyz.FIO, BankSvyz.ADR, BankSvyz.KV, BankSvyz.LIFT, BankSvyz.MUSOR, BankSvyz.SELEN, BankSvyz.GVoda, BankSvyz.Otopl, BankSvyz.HVoda, BankSvyz.SSLIV, BankSvyz.Key, BankSvyz.DATA, BankSvyz.OLDNUM, BankSvyz.PERIODOPL  FROM BankSvyz LEFT JOIN MainOccupant ON BankSvyz.LSCHET = MainOccupant.OLDNUM WHERE (((MainOccupant.OLDNUM) Is Null))"), Mconn, adOpenKeyset, adLockPessimistic

Set fg1.DataSource = rsNul


fg1.ColWidth(0) = 250






BtnEnh2.Visible = False
BtnEnh3.Visible = True

RsSet.Open ("Settings"), Mconn
Neo = RsSet("Neo")
RsSet.Close

RsSet.Open ("SELECT MainOccupant.Numer, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.КОД,MainOccupant.kv_num FROM KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom"), Mconn



End If




If Shag = 1 And fg1.Rows > 1 Then
Set rsNas = New ADODB.Recordset

rsNas.Open ("SELECT BankNastr.Код, BankNastr.ReestrPole, BankNastr.NachCod, nachisleniy.КодKategor, nachisleniy.Kategor, nachisleniy.Naim, nachisleniy.Formula, nachisleniy.Tip, nachisleniy.SchetZ, nachisleniy.FormulaB FROM BankNastr LEFT JOIN nachisleniy ON BankNastr.NachCod = nachisleniy.Kod"), Mconn





rsDoc.Open ("SELECT Doc.Cod,Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button, Doc.Dom , Doc.RealData FROM Doc"), Mconn, adOpenKeyset, adLockPessimistic

rsNas.MoveFirst
Do While Not rsNas.EOF

For rw = 1 To fg1.Rows - 1

If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
For fgcol = 1 To fg1.Cols - 1
If UCase(rsNas("ReestrPole")) = UCase(fg1.TextMatrix(0, fgcol)) And Val(fg1.TextMatrix(rw, fgcol)) <> 0 Then
rsDoc.AddNew
rsDoc("Cod") = Cod
rsDoc("DataR") = fg1.TextMatrix(rw, 20)
rsDoc("KodN") = rsNas("NachCod")
rsDoc("Summa") = Val(fg1.TextMatrix(rw, fgcol))
rsDoc("NameN") = rsNas("Naim")
rsDoc("KodKv") = fg1.TextMatrix(rw, 21)
rsDoc("NameKv") = fg1.TextMatrix(rw, 2) + " " + fg1.TextMatrix(rw, 3) + " " + fg1.TextMatrix(rw, 4)
rsDoc("Stst") = 0
'rsDoc("Com") = "Доб.реестр банка " + Reestr

rsDoc("Com") = "Доб.реестр банка " + Reestr + " опл. за " + fg1.TextMatrix(rw, 22)

rsDoc("Tip") = rsNas("Tip")
rsDoc("Dom") = fg1.TextMatrix(rw, 5)
'rsDoc("Realdata") = 0
rsDoc.UpdateBatch
End If
Next fgcol
rsDoc.UpdateBatch
End If
Next rw
rsNas.MoveNext
Loop

Set rsReestr = New ADODB.Recordset
rsReestr.Open ("Bank"), Mconn, adOpenKeyset, adLockPessimistic

For rw = 1 To fg1.Rows - 1
If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
DItem = fg1.TextMatrix(rw, 19)
fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked

rsReestr.MoveFirst
Do While Not rsReestr.EOF
If rsReestr("Key") = DItem Then
rsReestr.Delete
rsReestr.UpdateBatch
End If
rsReestr.MoveNext
Loop
End If
Next rw
rsDoc.Close
End If


rsDoobl.Requery



'Set FG1.DataSource = rsDoobl

End Sub

Private Sub BtnEnh3_Click()

'MsgBox Shag

rsDoc.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button, Doc.Dom FROM Doc"), Mconn, adOpenKeyset, adLockPessimistic

If Shag = 2 Then
If rsNas.State = adStateClosed Then
rsNas.Open ("SELECT BankNastr.Код, BankNastr.ReestrPole, BankNastr.NachCod, nachisleniy.КодKategor, nachisleniy.Kategor, nachisleniy.Naim, nachisleniy.Formula, nachisleniy.Tip, nachisleniy.SchetZ, nachisleniy.FormulaB FROM BankNastr LEFT JOIN nachisleniy ON BankNastr.NachCod = nachisleniy.Kod"), Mconn
End If


rsNas.MoveFirst
Do While Not rsNas.EOF
For rw = 1 To fg1.Rows - 1
If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
For fgcol = 1 To fg1.Cols - 1
If UCase(rsNas("ReestrPole")) = UCase(fg1.TextMatrix(0, fgcol)) And Val(fg1.TextMatrix(rw, fgcol)) <> 0 Then
rsDoc.AddNew
rsDoc("Cod") = Cod
rsDoc("DataR") = fg1.TextMatrix(rw, 20)
rsDoc("KodN") = rsNas("NachCod")
rsDoc("Summa") = Val(fg1.TextMatrix(rw, fgcol))
rsDoc("NameN") = rsNas("Naim")
rsDoc("KodKv") = fg1.TextMatrix(rw, 1)
rsDoc("NameKv") = fg1.TextMatrix(rw, 2) + " " + fg1.TextMatrix(rw, 3) + " " + fg1.TextMatrix(rw, 4)
rsDoc("Stst") = 0
'rsDoc("Com") = "Доб.реестр банка " + Reestr

'rsDoc("Com") = "Доб.реестр банка " + Reestr + " опл. за " + fg1.TextMatrix(Rw, 22)

rsDoc("Com") = "Доб.реестр банка " + Reestr + " опл. за " + fg1.TextMatrix(rw, 22)
rsDoc("Tip") = rsNas("Tip")
rsDoc("Dom") = fg1.TextMatrix(rw, 5)

rsDoc.UpdateBatch
End If
Next fgcol
rsDoc.UpdateBatch
End If
Next rw
rsNas.MoveNext
Loop
'rsDoc.Close


'rsReestr.Open ("Bank"), mconn, adOpenKeyset, adLockPessimistic

If rsReestr.State = adStateClosed Then rsReestr.Open ("Bank"), Mconn, adOpenKeyset, adLockPessimistic

rsReestr.Requery

For rw = 1 To fg1.Rows - 1
If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
DItem = fg1.TextMatrix(rw, 19)
fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked

rsReestr.MoveFirst
Do While Not rsReestr.EOF
If rsReestr("Key") = DItem Then
rsReestr.Delete
rsReestr.UpdateBatch
End If
rsReestr.MoveNext
Loop
End If
Next rw
rsNul.Requery
BtnEnh3.Caption = "Далее>>"


End If

fg1.Refresh


'********************************
'********************************
'********************************
If Shag = 4 Then
Set rsEnd = New ADODB.Recordset
rsEnd.Open ("SELECT BankSvyz.OLDNUM, BankSvyz.FAM, BankSvyz.IM, BankSvyz.OT, BankSvyz.КОД, BankSvyz.NAIM_KLS, BankSvyz.Num, BankSvyz.LSCHET, BankSvyz.FIO, BankSvyz.ADR, BankSvyz.KV, BankSvyz.LIFT, BankSvyz.MUSOR, BankSvyz.SELEN, BankSvyz.GVoda, BankSvyz.Otopl, BankSvyz.HVoda, BankSvyz.SSLIV, BankSvyz.Key, BankSvyz.DATA, BankSvyz.Numer, BankSvyz.PERIODOPL  From BankSvyz ORDER BY BankSvyz.OLDNUM, BankSvyz.KV, BankSvyz.LIFT, BankSvyz.MUSOR, BankSvyz.SELEN, BankSvyz.GVoda, BankSvyz.Otopl, BankSvyz.HVoda, BankSvyz.SSLIV, BankSvyz.PERIODOPL"), Mconn

Set fg1.DataSource = rsEnd


fg1.ColWidth(0) = 250

For rw = 1 To fg1.Rows - 1
fg1.Cell(flexcpChecked, rw, 0) = flexChecked
Next rw


For rw = 1 To fg1.Rows - 1
If Compare(UCase(fg1.TextMatrix(rw, 9)), UCase(fg1.TextMatrix(rw, 2)), 5) < 0.7 Then
fg1.Cell(flexcpBackColor, rw, 1, fg1.Rows - 1, fg1.Cols - 1) = RGB(200, 255, 100)
fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked
Else
fg1.Cell(flexcpBackColor, rw, 1, fg1.Rows - 1, fg1.Cols - 1) = vbWhite
End If
Next rw


Msg.Show
Msg.Label1.Caption = "Следующий список, показывает квартиросъемщиков, при идентификации номеров которых, не возникло проблем."
Msg.Label1.Caption = Msg.Label1.Caption + vbNewLine + "Все лиц.счета включенные в список отмечены, и готовы к разноске. Если вы не согласны с данным списком, то снимите отметку, и разнесите эту позицыю вручную."
Shag = 5
BtnEnh3.Visible = False
BtnEnh4.Visible = True

BtnEnh4.Caption = "Разнести"
End If




'********************************
'********************************
'********************************


'MSG.Show
'MSG.Label1.Caption = "Ок"

'Unload Me

'End If
rsDoc.Close

fg1.Refresh

If fg1.Rows <= 1 Then
BtnEnh3.Caption = "Далее>>"
Shag = 4
End If
End Sub

Private Sub BtnEnh4_1_Click()

End Sub

Private Sub BtnEnh4_Click()


If Shag = 5 Then

rsDoc.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button, Doc.Dom FROM Doc"), Mconn, adOpenKeyset, adLockPessimistic

rsNas.MoveFirst
Do While Not rsNas.EOF

For rw = 1 To fg1.Rows - 1

If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
For fgcol = 1 To fg1.Cols - 1
If UCase(rsNas("ReestrPole")) = UCase(fg1.TextMatrix(0, fgcol)) And Val(fg1.TextMatrix(rw, fgcol)) <> 0 Then
rsDoc.AddNew
rsDoc("Cod") = Cod
rsDoc("DataR") = fg1.TextMatrix(rw, 20)
rsDoc("KodN") = rsNas("NachCod")
rsDoc("Summa") = Val(fg1.TextMatrix(rw, fgcol))
rsDoc("NameN") = rsNas("Naim")
rsDoc("KodKv") = fg1.TextMatrix(rw, 21)
rsDoc("NameKv") = fg1.TextMatrix(rw, 2) + " " + fg1.TextMatrix(rw, 3) + " " + fg1.TextMatrix(rw, 4)
rsDoc("Stst") = 0
'rsDoc("Com") = "Доб.реестр банка " + Reestr

rsDoc("Com") = "Доб.реестр банка " + Reestr + " опл. за " + fg1.TextMatrix(rw, 22)
rsDoc("Tip") = rsNas("Tip")
rsDoc("Dom") = fg1.TextMatrix(rw, 5)

rsDoc.UpdateBatch
End If
Next fgcol
rsDoc.UpdateBatch
End If
Next rw
rsNas.MoveNext
Loop

rsReestr.Requery

For rw = 1 To fg1.Rows - 1
If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
DItem = fg1.TextMatrix(rw, 19)
fg1.Cell(flexcpChecked, rw, 0) = flexUnchecked

rsReestr.MoveFirst
Do While Not rsReestr.EOF
If rsReestr("Key") = DItem Then
rsReestr.Delete
rsReestr.UpdateBatch
End If
rsReestr.MoveNext
Loop
End If
Next rw

End If

rsEnd.Requery

BtnEnh4.Caption = "Далее >>>"



If fg1.Rows <= 1 Then
Msg.Show


If rsDoc.State = adStateClosed Then
rsDoc.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button, Doc.Dom FROM Doc"), Mconn, adOpenKeyset, adLockPessimistic
End If

rsDoc.MoveFirst

Do While Not rsDoc.EOF
If rsDoc("Cod") = Cod Then s = s + rsDoc("Summa")
rsDoc.MoveNext
Loop

Msg.Label1.Caption = "Данные добавлены в реестр документов оплаты №" + Str(Cod) + vbNewLine + "Сумма реестра банка=" + Str(SummI) + vbNewLine + "Разнесенная сумма=" + Str(s) + vbNewLine + "Отклонение=" + Str(Round(s - SummI, 2))
Unload Me
End If

rsDoc.Close
End Sub

Private Sub Command3_1_Click()

End Sub

Private Sub Command3_Click()
Pod.Show
Pod.Label1 = "Подождите идет экспорт данных в XL"

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

fg1.Subtotal flexSTClear
For I = 250 To 500
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

fg1.DataRefresh
For I = 500 To 750
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

ВыводВExel
For I = 750 To 1000
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

Unload Pod
End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Shag = 0 Then Проверка

If Shag = 2 And Col = 1 Then



RsSet.MoveFirst
Do While Not RsSet.EOF
If fg1.TextMatrix(Row, 1) = "" Then Exit Sub
If Trim(Str(RsSet("Numer"))) = Trim(Str(fg1.TextMatrix(Row, 1))) Then

If RsSet("Fam") <> "" Then fg1.TextMatrix(Row, 2) = RsSet("fam")
If RsSet("Im") <> "" Then fg1.TextMatrix(Row, 3) = RsSet("Im")
If RsSet("Ot") <> "" Then fg1.TextMatrix(Row, 4) = RsSet("Ot")
If RsSet("Код") <> "" Then fg1.TextMatrix(Row, 5) = RsSet("Код")
If RsSet("NAIM_KLS") <> "" Then fg1.TextMatrix(Row, 6) = RsSet("NAIM_KLS")
If RsSet("Num") <> "" Then fg1.TextMatrix(Row, 7) = RsSet("Num")
If RsSet("oldNum") <> "" Then fg1.TextMatrix(Row, 21) = RsSet("Oldnum")
Exit Do
End If
RsSet.MoveNext
Loop

fg1.AutoResize = True
fg1.Refresh
BtnEnh3.Caption = "Разнести"

For rw = 1 To fg1.Rows - 1
If fg1.TextMatrix(rw, 1) <> "" Then
fg1.Cell(flexcpChecked, rw, 0) = flexChecked
End If
Next rw

End If

'MsgBox Round(Val(Replace(FG1.TextMatrix(Rw, 11), ",", ".")), 2) - Round(Val(Replace(FG1.TextMatrix(Rw, 12), ",", ".")), 2)
'MsgBox Round(Val(Replace(FG1.TextMatrix(FG1.Row, 11), ",", ".")), 2) - Round(Val(Replace(FG1.TextMatrix(FG1.Row, 12), ",", ".")), 2)
'MsgBox Round(Val(Replace(FG1.TextMatrix(FG1.Row, 12), ",", ".")), 2)
'FG1.TextMatrix(Rw, 0) = Round(Round(Val(Replace(FG1.TextMatrix(Rw, 5), ",", ".")), 2) - (Round(Val(Replace(FG1.TextMatrix(Rw, 11), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 12), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 13), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 14), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 15), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 16), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 17), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 18), ",", ".")), 2)), 2)
End Sub

Private Sub FG1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If Shag = 2 And Col = 1 Then
fg1.ComboSearch = flexCmbSearchCombos
'MsgBox FG1.TextMatrix(Row, 9)
Comb = "#" + Neo + ";" + "Неопознанные суммы" + "|"
RsSet.MoveFirst
Do While Not RsSet.EOF
If RsSet("fam") <> "" Then

If Compare(UCase(fg1.TextMatrix(Row, 9)), UCase(RsSet("fam")), 5) > 0.5 Then
Comb = Comb + "#" + Str(RsSet("Numer")) + ";" + RsSet("fam") + " " + RsSet("Im") + " " + RsSet("Ot") + " " + RsSet("NAIM_KLS") + " кв.№" + RsSet("kv_num") + "|"
End If

End If
RsSet.MoveNext
Loop

fg1.ColComboList(1) = Comb



End If
If Col = 5 And fg1.TextMatrix(Row, Col) <> "" Then
If MsgBox("Общую сумму платежа править нельзя! Если Вы исправление сумму, то ОБЩАЯ СУММА ПО РЕЕСТРУ изменится, и не будет соответствовыть СУММЕ ВЫПИСКИ БАНКА за текущую дату" + vbNewLine + "ИСПРАВИТЬ?", vbYesNo) = vbNo Then Cancel = True
End If
End Sub

Private Sub Form_Load()
'MsgBox NewN

Fn = BankImport.File1.FileName
s = 0
Shag = 0
Beg = True
BankImport.Hide
ReestrDoc.Hide

Pod.Show

Pod.ProgressBar1.min = 1
BtnEnh2.Visible = False
BtnEnh3.Visible = False
BtnEnh4.Visible = False


ErrorStst = 0
'MakeWindow Me, True
fg1.Width = Me.Width / 15.40107
fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
' Command3.Top = Image1.Top
 
fg1.Sort = flexSortStringAscending

Set DBFConn1 = New ADODB.Connection
Set AcessConn = New ADODB.Connection

If BankImport.File1.FileName <> "" Then

DBFConn1.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы dBASE;Initial Catalog=" + BankImport.File1.Path

'DBFConn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + BankImport.File1.Path + "; Extended Properties=dBASE IV;User ID=Admin;Password=;"

Set RsSet = New ADODB.Recordset
Set dbfRs = New ADODB.Recordset
Set AccessRs = New ADODB.Recordset
Set rsDoc = New ADODB.Recordset
Set rsDocReestr = New ADODB.Recordset


Set rsReestr = New ADODB.Recordset
Set rsNas = New ADODB.Recordset




Pod.Refresh

'**************************************************************
'Временный укороченный файл
BankImport.File1.FileName = "tmp" + Right(BankImport.File1.FileName, 7)
For_Dell = BankImport.File1.Path + "/tmp" + Right(BankImport.File1.FileName, 7)

dbfRs.Open ("select * from " + BankImport.File1.FileName), DBFConn1
TSG = False
Dim myFld As ADODB.Field



' Проверка на БАНКЛИФТЫ и ЕРЦ
' Если количество полей Z=15 то это файл банка для лифтов если Z=13 то файл ЕРЦ
z = 0
For Each myFld In dbfRs.Fields
   ' If myFld.name = colname Then
    '    Old = False
     '   Exit For
    'End If
z = z + 1
Next

'***************
z = 17


'MsgBox (z)
'If z = 15 Then MainForm.LiftFile = True Else MainForm.Lift = False
If z = 18 Or z = 15 Then MainForm.LiftFile = True Else MainForm.Lift = False

If z = 13 Then MainForm.ErcFile = True Else MainForm.erc = False


colname = "SOPL"
Old = True
For Each myFld In dbfRs.Fields
    If myFld.name = colname Then
        Old = False
        Exit For
    End If
Next

colname = "NEWNUM"

NewN = False
For Each myFld In dbfRs.Fields
    If myFld.name = colname Then
        NewN = True
        Exit For
    End If
Next


'colname = "KPOKEL1"


For Each myFld In dbfRs.Fields
   If myFld.name = colname Then
      TSG = True
     Exit For
 End If
Next









If TSG Then
'MsgBox "TSG"
'TSGBank myFld
End If


dbfRs.Close
'****************************************************************

If MainForm.Old = True Then dbfRs.Open (BankImport.File1.FileName), DBFConn1, adOpenKeyset, adLockBatchOptimistic

If MainForm.Old = False And NewN = False Then
On Error GoTo nnn

dbfRs.Open ("SELECT KFOSB, TYPE, DATE as data, NUMBER, SOPL as Summa, SOPL_BK, FIO, ADR, LSCHET, PERIODOPL , SKOMM, SLIFT, SMUSOR, SELEN, SGVS, STEPLO, SHVODA, SSLIV, NUMPLP as PLNOM, DPLP as PLDATE, RSCHET as NRS FROM " + BankImport.File1.FileName), DBFConn1, adOpenKeyset, adLockBatchOptimistic

nnn:


'If Err.Number = -2147467259 Or Err.Number = 3021 Or Err.Number = 13 Or Err.Number = 0 Then
'Err.Clear
'Exit Sub
'Else

If Err.Number <> 0 Then
If Err.Number = -2147217904 Then

MsgBox Err.Description + " Ошибка №" + Str(Err.Number) + "   Попробую выполнить импорт в эталонный файл, в соответствии с настройками таблицы <<bank_sql>>"
Err.Clear
Me.BankImp

End If
End If



End If


If MainForm.Old = False And NewN = True Then


On Error GoTo nnn

dbfRs.Open ("SELECT KFOSB, TYPE, DATE as data, NUMBER, SOPL as Summa, SOPL_BK, FIO, ADR, LSCHET, PERIODOPL , SKOMM, SLIFT, SMUSOR, SELEN, SGVS, STEPLO, SHVODA, SSLIV, NUMPLP as PLNOM, DPLP as PLDATE, RSCHET as NRS, newnum FROM " + BankImport.File1.FileName), DBFConn1, adOpenKeyset, adLockBatchOptimistic
End If


'dbfRs.Open ("SELECT KFOSB, TYPE, NUMBER, FIO FROM " + BankImport.File1.FileName), DBFConn1, adOpenKeyset, adLockBatchOptimistic
lblTitle = "Импорт оплаты из банка. Файл > " + BankImport.File1.FileName
'F = BankImport.File1.FileName
Label1.Caption = "Просмотр файла >" + BankImport.File1.FileName + ". Для продолжения нажмите <<Далее>>"

Set fg1.DataSource = dbfRs

'MsgBox FG1.Cols
'MsgBox FG1.Rows
Pod.ProgressBar1.Max = (fg1.Cols * fg1.Rows) * 2
' Заменяем разделитель разрядов
For Cl = 1 To fg1.Cols - 1
For rw = 1 To fg1.Rows - 1

DoEvents

Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1

fg1.TextMatrix(rw, Cl) = Replace(fg1.TextMatrix(rw, Cl), ".", ",")
Next rw
Next Cl


fg1.ColHidden(2) = True
fg1.ColHidden(3) = True
fg1.ColHidden(4) = True



If MainForm.Old = False Then Проверка
SummI = 0
For rw = 1 To fg1.Rows - 1
SummI = SummI + Round(Val(Replace(fg1.TextMatrix(rw, 5), ",", ".")), 2)
Next

lblTitle = lblTitle + "На сумму > " + Str(SummI)


Unload Pod
Beg = False
Unload BankImport
Else
Unload Pod
Unload BankImport
'MsgBox "Вы не выбрали файл для импорта!"
lblTitle = "!! Файл не указан !! "

End If




End Sub

Private Sub Form_Unload(Cancel As Integer)
ReestrDoc.Enabled = True
MainMenu.Enabled = True
End Sub

Private Sub Image1_Click()
Unload Me
ReestrDoc.Enabled = True
End Sub
Sub ВыводВExel()
   Const НачСтрока = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim I As Long, J As Long, K As Long, rДанные As String
   Dim v As Variant
   
   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
   
   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(fg1.Cols - 1) & fg1.Rows + НачСтрока
   ReDim v(fg1.Rows, fg1.Cols) 'Забыл указать
   
   If fg1.Rows > 0 Then
            For co = 1 To fg1.Cols - 1
         For rw = 0 To fg1.Rows - 1

             v(rw, co) = fg1.TextMatrix(rw, co)
             
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
Form2.Label1 = "   Импорт оплаты коммунальных услуг, из файла данных предоставленных банком. Это окно предназначено для предварительного просмотра импортируемых данных." + vbNewLine + "   Кроме того, Вы можете перенести данные в XL для дальнейшей распечатки"


Form2.Show
End Sub

Private Sub imgTitleMain_Click()
ChangeState Me
End Sub

Private Sub lblTitle_Click()
ChangeState Me
End Sub
Private Sub Form_Resize()
fg1.Width = Me.Width / 15.40107
   fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
BtnEnh1.Top = Image1.Top
BtnEnh2.Top = Image1.Top
BtnEnh3.Top = Image1.Top
'Command1.Top = Image1.Top
'Command5.Top = Image1.Top
'Command9.Top = Image1.Top
'Command2.Top = Image1.Top
'Command11.Top = Image1.Top

Command3.Left = Image1.Left + Image1.Width
'BtnEnh1.Top = Image1.Left + Image1.Width + Command3.Width
'Command4.Left = Image1.Left + Image1.Width + Command3.Width
'Command5.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width
'Command1.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width
'Command9.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width + Command1.Width
'Command2.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width + Command1.Width + Command9.Width
'Command11.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width + Command1.Width + Command9.Width + Command2.Width
End Sub
Private Sub Проверка()

' Первоначальные цвета
fg1.Cell(flexcpForeColor, 1, 1, fg1.Rows - 1, fg1.Cols - 1) = vbBlack
fg1.Cell(flexcpFontBold, 1, 1, fg1.Rows - 1, fg1.Cols - 1) = False
fg1.Cell(flexcpBackColor, 1, 1, fg1.Rows - 1, fg1.Cols - 1) = vbWhite
fg1.Cell(flexcpBackColor, 1, 0, fg1.Rows - 1, 0) = &H8000000F


ErrorStst = 0


' Выделение ошибок цветом

For rw = 1 To fg1.Rows - 1

'MsgBox FG1.TextMatrix(Rw, 22)

For Cl = 1 To fg1.Cols - 1

If Beg = True Then Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
'FG1.TextMatrix(Rw, Cl) = Replace(FG1.TextMatrix(Rw, Cl), ",", ".")

If fg1.TextMatrix(rw, Cl) = "" Then
If Cl <> 5 And Cl <> 7 And Cl <> 9 Then
fg1.TextMatrix(rw, Cl) = 0
fg1.Cell(flexcpForeColor, rw, Cl, rw, Cl) = vbRed
End If
If NewN = True Then
If (Cl = 5 Or Cl = 7 Or Cl = 9) And Len(fg1.TextMatrix(rw, 22)) <> 12 Then
ErrorStst = ErrorStst + 1
fg1.Cell(flexcpBackColor, rw, Cl, rw, Cl) = vbMagenta
fg1.Cell(flexcpForeColor, rw, Cl, rw, Cl) = vbWhite
fg1.Cell(flexcpFontBold, rw, Cl, rw, Cl) = True
fg1.Cell(flexcpBackColor, rw, 0, rw, 0) = vbRed
End If
End If


If NewN = False Then
If (Cl = 5 Or Cl = 7 Or Cl = 9) Then
ErrorStst = ErrorStst + 1
fg1.Cell(flexcpBackColor, rw, Cl, rw, Cl) = vbMagenta
fg1.Cell(flexcpForeColor, rw, Cl, rw, Cl) = vbWhite
fg1.Cell(flexcpFontBold, rw, Cl, rw, Cl) = True
fg1.Cell(flexcpBackColor, rw, 0, rw, 0) = vbRed
End If
End If





End If




Next Cl

Ssum = Round(Val(Replace(fg1.TextMatrix(rw, 5), ",", ".")), 2)
Kv = Round(Val(Replace(fg1.TextMatrix(rw, 11), ",", ".")), 2)
Lift = Round(Val(Replace(fg1.TextMatrix(rw, 12), ",", ".")), 2)
Musor = Round(Val(Replace(fg1.TextMatrix(rw, 13), ",", ".")), 2)
El = Round(Val(Replace(fg1.TextMatrix(rw, 14), ",", ".")), 2)
Gv = Round(Val(Replace(fg1.TextMatrix(rw, 15), ",", ".")), 2)
Otopl = Round(Val(Replace(fg1.TextMatrix(rw, 16), ",", ".")), 2)
Hv = Round(Val(Replace(fg1.TextMatrix(rw, 17), ",", ".")), 2)
Sliv = Round(Val(Replace(fg1.TextMatrix(rw, 18), ",", ".")), 2)


'MsgBox Round(Val(Replace(FG1.TextMatrix(Rw, 11), ",", ".")), 2)

If Round(Ssum, 2) <> Round(Kv + Lift + Musor + El + Gv + Otopl + Hv + Sliv, 2) Then

ErrorStst = ErrorStst + 1
'MsgBox Str(Kv + Lift + Musor + El + Gv + Otopl + Hv + Sliv)
'MsgBox Round(Round(Val(Replace(FG1.TextMatrix(Rw, 11), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 12), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 13), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 14), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 15), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 16), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 17), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 18), ",", ".")), 2), 2)

'Val(Replace(Str(FG1.TextMatrix(Rw, 11)), ".", ",")) Then
'+ Val(Str(FG1.TextMatrix(Rw, 12))) + Val(Str(FG1.TextMatrix(Rw, 13))) + Val(Str(FG1.TextMatrix(Rw, 14))) + Val(Str(FG1.TextMatrix(Rw, 15))) + Val(Str(FG1.TextMatrix(Rw, 16))) + Val(Str(FG1.TextMatrix(Rw, 17))) + Val(Str(FG1.TextMatrix(Rw, 18))) Then

fg1.Cell(flexcpBackColor, rw, 0, rw, fg1.Cols - 1) = &H80000018
fg1.Cell(flexcpBackColor, rw, 5, rw, 5) = vbYellow
fg1.Cell(flexcpBackColor, rw, 0, rw, 0) = vbYellow




fg1.TextMatrix(rw, 0) = Round(Ssum - (Kv + Lift + Musor + El + Gv + Otopl + Hv + Sliv), 2)
Else
fg1.TextMatrix(rw, 0) = 0
End If

Next rw

End Sub


Public Sub BankImp() ' Процедура построения запроса для импорта из банка файлов любых форматов
Set rsBank = New ADODB.Recordset
Set rsBank.ActiveConnection = Mconn
rsBank.Open ("SELECT bank_sql.KFOSB, bank_sql.TYPE, bank_sql.DATE, bank_sql.NUMBER, bank_sql.SOPL, bank_sql.FIO, bank_sql.ADR, bank_sql.LSCHET, bank_sql.PERIODOPL, bank_sql.SKOMM, bank_sql.SLIFT, bank_sql.SMUSOR, bank_sql.NUMPLP, bank_sql.DPLP, bank_sql.RSCHET, bank_sql.OPER_OPL, bank_sql.SOPL_BK, bank_sql.SELEN, bank_sql.SGVS, bank_sql.STEPLO, bank_sql.SHVODA, bank_sql.SSLIV, bank_sql.SGAZ, bank_sql.SDOPL, bank_sql.NEWNUM, bank_sql.PLATUSL, bank_sql.NPOKEL1, bank_sql.KPOKEL1, bank_sql.NPOKHV1, bank_sql.KPOKHV1, bank_sql.NPOKHV2, bank_sql.KPOKHV2, bank_sql.ADRES, bank_sql.NAZN, bank_sql.NAZN1, bank_sql.coment FROM bank_sql")





'FileCopy App.Path + "/Dbf/BETEM.DBF", App.Path + "/dbf/" + "BET.DBF"

FileCopy App.Path + "\Dbf\BETEM.DBF", App.Path + "\dbf\" + "BET.DBF"

dbfRs.Open ("SELECT * FROM " + BankImport.File1.FileName), DBFConn1, adOpenKeyset, adLockBatchOptimistic

КоннектDBF
Set rsDbfBank = New ADODB.Recordset



rsDbfBank.Open ("select * from BET.DBF"), DBFConn, adOpenKeyset, adLockBatchOptimistic

'Цикл по настройкам экспорта
rsBank.MoveFirst






Do While Not rsBank.EOF





Me.Caption = "Пробую настройку экспорта " + rsBank("coment")
ReestrDoc.Caption = "Пробую настройку экспорта " + rsBank("coment")





dbfRs.MoveFirst


On Error GoTo Ec



Do While Not dbfRs.EOF


On Error GoTo Ec



DoEvents
Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 0.1
Pod.Label1 = st + " Пробую настройку экспорта " + rsBank("coment")

rsDbfBank.AddNew



st = Trim(rsBank("KFOSB"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("KFOSB") = "0" Else rsDbfBank("KFOSB") = rsBank("ST")

st = Trim(rsBank("TYPE"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("TYPE") = "0" Else rsDbfBank("TYPE") = dbfRs(st)

st = Trim(rsBank("Date"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("Date") = "0" Else rsDbfBank("Date") = dbfRs(st)



st = Trim(rsBank("Number"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh

'Проверка на непустое значение номера лиц.счета

'If dbfRs(st) = "0050761003" Then MsgBox (dbfRs(st))

If dbfRs(st) = Null Then

dbfRs(st) = "0"

End If

If st = "0" Then rsDbfBank("Number") = "0" Else rsDbfBank("Number") = dbfRs(st)

st = Trim(rsBank("SOPL"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SOPL") = "0" Else rsDbfBank("SOPL") = dbfRs(st)

st = Trim(rsBank("FIO"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("FIO") = "0" Else rsDbfBank("FIO") = dbfRs(st)

st = Trim(rsBank("Adr"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("Adr") = "0" Else rsDbfBank("Adr") = dbfRs(st)





st = Trim(rsBank("LSCHET"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
Bn = dbfRs(st)

Do While Len(Bn) < 12
Bn = "0" + Bn
Loop

If st = "0" Then rsDbfBank("LSCHET") = "0" Else rsDbfBank("LSCHET") = Bn

st = Trim(rsBank("PERIODOPL"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("PERIODOPL") = "0" Else rsDbfBank("PERIODOPL") = dbfRs(st)

st = Trim(rsBank("SKOMM"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SKOMM") = "0" Else rsDbfBank("SKOMM") = dbfRs(st)

st = Trim(rsBank("SLIFT"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SLIFT") = "0" Else rsDbfBank("SLIFT") = dbfRs(st)


st = Trim(rsBank("SMUSOR"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SMUSOR") = "0" Else rsDbfBank("SMUSOR") = dbfRs(st)


st = Trim(rsBank("NUMPLP"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("NUMPLP") = "0" Else rsDbfBank("NUMPLP") = dbfRs(st)


st = Trim(rsBank("DPLP"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("DPLP") = Date Else rsDbfBank("DPLP") = dbfRs(st)

st = Trim(rsBank("RSCHET"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("RSCHET") = "0" Else rsDbfBank("RSCHET") = dbfRs(st)

st = Trim(rsBank("OPER_OPL"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("OPER_OPL") = Date Else rsDbfBank("OPER_OPL") = dbfRs(st)

st = Trim(rsBank("SOPL_BK"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SOPL_BK") = "0" Else rsDbfBank("SOPL_BK") = dbfRs(st)

st = Trim(rsBank("SELEN"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SELEN") = "0" Else rsDbfBank("SELEN") = dbfRs(st)

st = Trim(rsBank("SGVS"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SGVS") = "0" Else rsDbfBank("SGVS") = dbfRs(st)

st = Trim(rsBank("STEPLO"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("STEPLO") = "0" Else rsDbfBank("STEPLO") = dbfRs(st)

st = Trim(rsBank("SHVODA"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SHVODA") = "0" Else rsDbfBank("SHVODA") = dbfRs(st)

st = Trim(rsBank("SSLIV"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SSLIV") = "0" Else rsDbfBank("SSLIV") = dbfRs(st)

st = Trim(rsBank("SGAZ"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SGAZ") = "0" Else rsDbfBank("SGAZ") = dbfRs(st)

st = Trim(rsBank("SDOPL"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("SDOPL") = "0" Else rsDbfBank("SDOPL") = dbfRs(st)

st = Trim(rsBank("NEWNUM"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh

'Дополняем нолями номера лиц счетов
Bn = dbfRs(st)

Do While Len(Bn) < 12
Bn = "0" + Bn
Loop

If st = "0" Then rsDbfBank("NEWNUM") = "0" Else rsDbfBank("NEWNUM") = Bn

st = Trim(rsBank("PLATUSL"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("PLATUSL") = "0" Else rsDbfBank("PLATUSL") = dbfRs(st)

st = Trim(rsBank("NAZN"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("NAZN") = "0" Else rsDbfBank("NAZN") = Right(dbfRs(st), 40)

st = Trim(rsBank("NAZN1"))
Pod.Label1.Caption = " Обрабатываю " + st
Pod.Label1.Refresh
If st = "0" Then rsDbfBank("NAZN1") = "0" Else rsDbfBank("NAZN1") = dbfRs(st)

rsDbfBank.Update
 

dbfRs.MoveNext






Loop


If dbfRs.EOF Then Exit Do

Ec:

If Err.Number <> 0 Then
'Do While Not rsBank.EOF
MsgBox ("Ошибка настроек:" + rsBank("coment") + "  прбую следующую. Номер ошибки:" + Err.Description + Str(Err.Number) + " след за ошибкой поле: " + st)
rsDbfBank.Delete adAffectCurrent
'FileCopy App.Path + "/Dbf/BETEM.DBF", App.Path + "/dbf/" + "BET.DBF"

' Меняем настройки экспорта
rsBank.MoveNext
MsgBox ("Применяю настройку " + rsBank("coment"))
'Loop


Err.Clear

End If



' Меняем настройки экспорта
'rsBank.MoveNext


'Конец цикла по настройкам экспорта
Loop

rsDbfBank.UpdateBatch


' открыли bank_sql с информацией о полях файла банка
' открыли bank_etalon.DBF для добавления
' и пробуем вставить данные из файла банка в bet.DBF

rsDbfBank.Close
Pod.Hide


FileCopy App.Path + "\Dbf\BET.DBF", BankImport.File1.Path + "\BET.DBF"
Fn = "bet.dbf"
BankImport.File1.FileName = "bet.dbf"





' Эту часть переделать

MainForm.Old = False
NewN = True

dbfRs.Close
'dbfRs.Open ("SELECT KFOSB, TYPE, DATE as data, NUMBER, SOPL as Summa, SOPL_BK, FIO, ADR, LSCHET, PERIODOPL , SKOMM, SLIFT, SMUSOR, SELEN, SGVS, STEPLO, SHVODA, SSLIV, NUMPLP as PLNOM, DPLP as PLDATE, RSCHET as NRS FROM bet.dbf"), DBFConn1, adOpenKeyset, adLockBatchOptimistic
'dbfRs.Open ("SELECT KFOSB, TYPE, DATE as data, NUMBER, SOPL as Summa, SOPL_BK, FIO, ADR, LSCHET, PERIODOPL , SKOMM, SLIFT, SMUSOR, SELEN, SGVS, STEPLO, SHVODA, SSLIV, NUMPLP as PLNOM, DPLP as PLDATE, RSCHET as NRS, newnum FROM bet.dbf"), DBFConn1, adOpenKeyset, adLockBatchOptimistic

'**********************************


End Sub

