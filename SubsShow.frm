VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form SubsShow 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8790
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11865
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "XL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Image1 
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
      TabIndex        =   3
      Top             =   8160
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   9495
      _cx             =   16748
      _cy             =   9975
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"SubsShow.frx":0000
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
      Height          =   195
      Left            =   0
      Picture         =   "SubsShow.frx":00E2
      Top             =   0
      Width           =   195
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
      Left            =   720
      TabIndex        =   1
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10890
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "SubsShow.frx":032C
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   360
      Picture         =   "SubsShow.frx":0A76
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "SubsShow.frx":11C0
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "SubsShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBFConn As ADODB.Connection
'Dim mconn As ADODB.Connection
Dim dbfRs As ADODB.Recordset
Dim AccessRs As ADODB.Recordset
Dim ErrorStst As Integer
Dim rsDoobl As ADODB.Recordset
Dim RsSet As ADODB.Recordset
Dim rsNul As ADODB.Recordset
Dim rsOk As ADODB.Recordset
Dim Shag As Integer
Dim rsDoc As ADODB.Recordset
Dim rsDocReestr As ADODB.Recordset
Dim Cod As Integer
Dim rsNas As ADODB.Recordset
Dim rsReestr As ADODB.Recordset
Dim Neo As String
Public Reestr As String
Dim DItem As String
Dim Clik As Integer
Dim s As Double
Dim Beg As Boolean
Dim Ssum As Double, Kv As Double, Musor As Double, Lift As Double, Otopl As Double, Gv As Double, Hv As Double, El As Double, Sliv As Double
Dim rsEnd As ADODB.Recordset
Dim SummI As Double
Dim Fn As String
Dim Old As Boolean






Private Sub BtnEnh4_Click()

End Sub

Private Sub Command3_Click()
Pod.Show
Pod.Label1 = "Подождите идет экспорт данных в XL"

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
 For j = 1 To 1000
    Next j
   Next

FG1.Subtotal flexSTClear
For I = 250 To 500
    Pod.ProgressBar1.Value = I
 For j = 1 To 1000
    Next j
   Next

FG1.DataRefresh
For I = 500 To 750
    Pod.ProgressBar1.Value = I
    
 For j = 1 To 1000
    Next j
   Next

ВыводВExel
For I = 750 To 1000
    Pod.ProgressBar1.Value = I
    
 For j = 1 To 1000
    Next j
   Next

Unload Pod
End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Shag = 0 Then Проверка

If Shag = 2 And Col = 1 Then



RsSet.MoveFirst
Do While Not RsSet.EOF
If FG1.TextMatrix(Row, 1) = "" Then Exit Sub
If Trim(Str(RsSet("Numer"))) = Trim(Str(FG1.TextMatrix(Row, 1))) Then

If RsSet("Fam") <> "" Then FG1.TextMatrix(Row, 2) = RsSet("fam")
If RsSet("Im") <> "" Then FG1.TextMatrix(Row, 3) = RsSet("Im")
If RsSet("Ot") <> "" Then FG1.TextMatrix(Row, 4) = RsSet("Ot")
If RsSet("Код") <> "" Then FG1.TextMatrix(Row, 5) = RsSet("Код")
If RsSet("NAIM_KLS") <> "" Then FG1.TextMatrix(Row, 6) = RsSet("NAIM_KLS")
If RsSet("Num") <> "" Then FG1.TextMatrix(Row, 7) = RsSet("Num")
If RsSet("oldNum") <> "" Then FG1.TextMatrix(Row, 21) = RsSet("Oldnum")
Exit Do
End If
RsSet.MoveNext
Loop

FG1.AutoResize = True
FG1.Refresh
BtnEnh3.Caption = "Разнести"

For rw = 1 To FG1.Rows - 1
If FG1.TextMatrix(rw, 1) <> "" Then
FG1.Cell(flexcpChecked, rw, 0) = flexChecked
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
FG1.ComboSearch = flexCmbSearchCombos
'MsgBox FG1.TextMatrix(Row, 9)
Comb = "#" + Neo + ";" + "Неопознанные суммы" + "|"
RsSet.MoveFirst
Do While Not RsSet.EOF
If RsSet("fam") <> "" Then

If Compare(UCase(FG1.TextMatrix(Row, 9)), UCase(RsSet("fam")), 5) > 0.5 Then
Comb = Comb + "#" + Str(RsSet("Numer")) + ";" + RsSet("fam") + " " + RsSet("Im") + " " + RsSet("Ot") + " " + RsSet("NAIM_KLS") + " кв.№" + RsSet("kv_num") + "|"
End If

End If
RsSet.MoveNext
Loop

FG1.ColComboList(1) = Comb



End If
If Col = 5 And FG1.TextMatrix(Row, Col) <> "" Then
If MsgBox("Общую сумму платежа править нельзя! Если Вы исправление сумму, то ОБЩАЯ СУММА ПО РЕЕСТРУ изменится, и не будет соответствовыть СУММЕ ВЫПИСКИ БАНКА за текущую дату" + vbNewLine + "ИСПРАВИТЬ?", vbYesNo) = vbNo Then Cancel = True
End If
End Sub

Private Sub Form_Load()
Shag = 0
Beg = True
Pod.Show
Pod.ProgressBar1.min = 1
BtnEnh2.Visible = False
BtnEnh3.Visible = False
BtnEnh4.Visible = False


ErrorStst = 0
MakeWindow Me, True
FG1.Width = Me.Width / 15.40107
FG1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
 Command3.Top = Image1.Top
 
FG1.Sort = flexSortStringAscending

Set DBFConn = New ADODB.Connection
Set AcessConn = New ADODB.Connection

If SubsImport.File1.FileName <> "" Then
'DBFConn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы dBASE;Initial Catalog=" + SubsImport.File1.Path

'DBFConn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы XLS;Initial Catalog=" + SubsImport.File1.Path
DBFConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + SubsImport.File1.Path + "\" + SubsImport.File1.FileName + ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""

Set RsSet = New ADODB.Recordset
Set dbfRs = New ADODB.Recordset
Set AccessRs = New ADODB.Recordset
Set rsDoc = New ADODB.Recordset
Set rsDocReestr = New ADODB.Recordset


Set rsReestr = New ADODB.Recordset
Set rsNas = New ADODB.Recordset




Pod.Refresh



dbfRs.Open ("SELECT * from [" + Mid(SubsImport.File1.FileName, 1, Len(SubsImport.File1.FileName) - 4) + "$] "), DBFConn




'dbfRs.Open ("[GEK2$]"), DBFConn
'dbfRs.Close
'lold:
'MsgBox Err.Number


'MsgBox Err.Description


'If Old = True Then dbfRs.Open (SubsImport.File1.FileName), DBFConn, adOpenKeyset, adLockBatchOptimistic

'If Old = False Then dbfRs.Open ("SELECT KFOSB, TYPE, DATE as data, NUMBER, SOPL as Summa, SOPL_BK, FIO, ADR, LSCHET, PERIODOPL , SKOMM, SLIFT, SMUSOR, SELEN, SGVS, STEPLO, SHVODA, SSLIV, NUMPLP as PLNOM, DPLP as PLDATE, RSCHET as NRS FROM " + SubsImport.File1.FileName), DBFConn, adOpenKeyset, adLockBatchOptimistic


'dbfRs.Open ("SELECT KFOSB, TYPE, NUMBER, FIO FROM " + subsImport.File1.FileName), DBFConn, adOpenKeyset, adLockBatchOptimistic
lblTitle = "Импорт оплаты из банка. Файл > " + SubsImport.File1.FileName
Label1.Caption = "Просмотр файла >" + SubsImport.File1.FileName + ". Для продолжения нажмите <<Далее>>"

Set FG1.DataSource = dbfRs


lblTitle = lblTitle + "На сумму > " + Str(SummI)


Unload Pod
Beg = False
Unload SubsImport
Else
Unload Pod
Unload SubsImport
MsgBox "Вы не выбрали файл для импорта!"
lblTitle = "!! Файл не указан !! "
End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
Sub ВыводВExel()
   Const НачСтрока = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim I As Long, j As Long, k As Long, rДанные As String
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
FG1.Width = Me.Width / 15.40107
   FG1.Height = Me.Height / 20
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
FG1.Cell(flexcpForeColor, 1, 1, FG1.Rows - 1, FG1.Cols - 1) = vbBlack
FG1.Cell(flexcpFontBold, 1, 1, FG1.Rows - 1, FG1.Cols - 1) = False
FG1.Cell(flexcpBackColor, 1, 1, FG1.Rows - 1, FG1.Cols - 1) = vbWhite
FG1.Cell(flexcpBackColor, 1, 0, FG1.Rows - 1, 0) = &H8000000F


ErrorStst = 0


' Выделение ошибок цветом

For rw = 1 To FG1.Rows - 1
For Cl = 1 To FG1.Cols - 1

If Beg = True Then Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
'FG1.TextMatrix(Rw, Cl) = Replace(FG1.TextMatrix(Rw, Cl), ",", ".")

If FG1.TextMatrix(rw, Cl) = "" Then
If Cl <> 5 And Cl <> 7 And Cl <> 9 Then
FG1.TextMatrix(rw, Cl) = 0
FG1.Cell(flexcpForeColor, rw, Cl, rw, Cl) = vbRed
End If

If Cl = 5 Or Cl = 7 Or Cl = 9 Then
ErrorStst = ErrorStst + 1
FG1.Cell(flexcpBackColor, rw, Cl, rw, Cl) = vbMagenta
FG1.Cell(flexcpForeColor, rw, Cl, rw, Cl) = vbWhite
FG1.Cell(flexcpFontBold, rw, Cl, rw, Cl) = True
FG1.Cell(flexcpBackColor, rw, 0, rw, 0) = vbRed
End If
End If




Next Cl

Ssum = Round(Val(Replace(FG1.TextMatrix(rw, 5), ",", ".")), 2)
Kv = Round(Val(Replace(FG1.TextMatrix(rw, 11), ",", ".")), 2)
Lift = Round(Val(Replace(FG1.TextMatrix(rw, 12), ",", ".")), 2)
Musor = Round(Val(Replace(FG1.TextMatrix(rw, 13), ",", ".")), 2)
El = Round(Val(Replace(FG1.TextMatrix(rw, 14), ",", ".")), 2)
Gv = Round(Val(Replace(FG1.TextMatrix(rw, 15), ",", ".")), 2)
Otopl = Round(Val(Replace(FG1.TextMatrix(rw, 16), ",", ".")), 2)
Hv = Round(Val(Replace(FG1.TextMatrix(rw, 17), ",", ".")), 2)
Sliv = Round(Val(Replace(FG1.TextMatrix(rw, 18), ",", ".")), 2)
'MsgBox Round(Val(Replace(FG1.TextMatrix(Rw, 11), ",", ".")), 2)

If Round(Ssum, 2) <> Round(Kv + Lift + Musor + El + Gv + Otopl + Hv + Sliv, 2) Then

ErrorStst = ErrorStst + 1
'MsgBox Str(Kv + Lift + Musor + El + Gv + Otopl + Hv + Sliv)
'MsgBox Round(Round(Val(Replace(FG1.TextMatrix(Rw, 11), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 12), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 13), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 14), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 15), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 16), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 17), ",", ".")), 2) + Round(Val(Replace(FG1.TextMatrix(Rw, 18), ",", ".")), 2), 2)

'Val(Replace(Str(FG1.TextMatrix(Rw, 11)), ".", ",")) Then
'+ Val(Str(FG1.TextMatrix(Rw, 12))) + Val(Str(FG1.TextMatrix(Rw, 13))) + Val(Str(FG1.TextMatrix(Rw, 14))) + Val(Str(FG1.TextMatrix(Rw, 15))) + Val(Str(FG1.TextMatrix(Rw, 16))) + Val(Str(FG1.TextMatrix(Rw, 17))) + Val(Str(FG1.TextMatrix(Rw, 18))) Then

FG1.Cell(flexcpBackColor, rw, 0, rw, FG1.Cols - 1) = &H80000018
FG1.Cell(flexcpBackColor, rw, 5, rw, 5) = vbYellow
FG1.Cell(flexcpBackColor, rw, 0, rw, 0) = vbYellow




FG1.TextMatrix(rw, 0) = Round(Ssum - (Kv + Lift + Musor + El + Gv + Otopl + Hv + Sliv), 2)
Else
FG1.TextMatrix(rw, 0) = 0
End If

Next rw

End Sub
