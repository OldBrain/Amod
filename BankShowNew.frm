VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form BankShowNew 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8796
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   11868
   ControlBox      =   0   'False
   Icon            =   "BankShowNew.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   733
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Далее >>>"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
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
      TabIndex        =   4
      Top             =   8160
      Width           =   1455
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
      FormatString    =   $"BankShowNew.frx":030A
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
      Picture         =   "BankShowNew.frx":03EC
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
      Picture         =   "BankShowNew.frx":0636
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BankShowNew.frx":0D80
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BankShowNew.frx":14CA
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "BankShowNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNewN As ADODB.Recordset
Dim rsDoc As ADODB.Recordset
Dim rsReestr As ADODB.Recordset
Dim rsNas As ADODB.Recordset





Private Sub BtnEnh1_Click()

Set rsDoc = New ADODB.Recordset
Set rsNas = New ADODB.Recordset
Set rsReestr = New ADODB.Recordset


If rsNas.State = adStateClosed Then
rsNas.Open ("SELECT BankNastr.Код, BankNastr.ReestrPole, BankNastr.NachCod, nachisleniy.КодKategor, nachisleniy.Kategor, nachisleniy.Naim, nachisleniy.Formula, nachisleniy.Tip, nachisleniy.SchetZ, nachisleniy.FormulaB FROM BankNastr LEFT JOIN nachisleniy ON BankNastr.NachCod = nachisleniy.Kod"), Mconn
End If


rsDoc.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button, Doc.Dom, Doc.Realdata, doc.plnom FROM Doc"), Mconn, adOpenKeyset, adLockPessimistic

If rsNas.RecordCount > 0 Then rsNas.MoveFirst

For rw = 1 To fg1.Rows - 1
fg1.Cell(flexcpChecked, rw, 0) = flexChecked
'flexChecked
Next rw


Do While Not rsNas.EOF

For rw = 1 To fg1.Rows - 1

If fg1.Cell(flexcpChecked, rw, 0) = flexChecked Then
For fgcol = 1 To fg1.Cols - 1
If UCase(rsNas("ReestrPole")) = UCase(fg1.TextMatrix(0, fgcol)) And Val(fg1.TextMatrix(rw, fgcol)) <> 0 Then
rsDoc.AddNew
rsDoc("Cod") = BankShow.Cod
'If fg1.TextMatrix(rw, 21) = 0 Then
rsDoc("DataR") = Date
'Else rsDoc("DataR") = fg1.TextMatrix(rw, 21)

rsDoc("KodN") = rsNas("NachCod")
rsDoc("Summa") = Val(fg1.TextMatrix(rw, fgcol))
rsDoc("NameN") = rsNas("Naim")
rsDoc("KodKv") = fg1.TextMatrix(rw, 22)

rsDoc("NameKv") = fg1.TextMatrix(rw, 3) + " " + fg1.TextMatrix(rw, 4) + " " + fg1.TextMatrix(rw, 5)
rsDoc("Stst") = 0
'rsDoc("Com") = "Доб.реестр банка " + Reestr

rsDoc("Com") = "р-р банка " + BankShow.Reestr + "п/п №" + fg1.TextMatrix(rw, 24) + "от " + fg1.TextMatrix(rw, 21) + " опл.за " + fg1.TextMatrix(rw, 23)
rsDoc("Tip") = rsNas("Tip")
rsDoc("Dom") = fg1.TextMatrix(rw, 7)
rsDoc("plnom") = fg1.TextMatrix(rw, 24)

'Оплата за период....
On Error GoTo prop

 

If (fg1.TextMatrix(rw, 23)) = "0" Then GoTo prop
'MsgBox (fg1.TextMatrix(rw, 21))
rsDoc("RealData") = Replace(Replace(Replace(fg1.TextMatrix(rw, 23), "-", "/"), ",", "/"), "00", "01")

prop:
If Err.Number <> 0 Then

rsDoc("RealData") = fg1.TextMatrix(rw, 21)
Err.Clear
End If
rsDoc.UpdateBatch
End If
Next fgcol
rsDoc.UpdateBatch
End If
Next rw
rsNas.MoveNext
Loop

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

If Err.Number = 0 And fg1.Rows <> 1 Then

rsNewN.Open ("bank"), Mconn, adOpenKeyset, adLockPessimistic
rsNewN.MoveFirst
Do While Not rsNewN.EOF
For rw = 1 To fg1.Rows - 1
If rsNewN("NewNum") = fg1.TextMatrix(rw, 1) Then
rsNewN.Delete
rsNewN.UpdateBatch
Exit For
End If
Next
rsNewN.MoveNext
Loop
rsNewN.Close
Else
If Err.Description <> "" Then MsgBox Err.Description
Err.Clear
End If


If rsDoc.State = adStateClosed Then
rsDoc.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button, Doc.Dom FROM Doc"), Mconn, adOpenKeyset, adLockPessimistic
End If

If rsDoc.RecordCount > 0 Then rsDoc.MoveFirst
BankShow.s = 0
Do While Not rsDoc.EOF
If rsDoc("Cod") = BankShow.Cod Then BankShow.s = BankShow.s + rsDoc("Summa")
rsDoc.MoveNext
Loop
rsDoc.Close



Unload Me
End Sub

Private Sub Command3__Click()

End Sub

Private Sub BtnEnh11_Click()

End Sub

Private Sub Command3_Click()


ВыводВExel


End Sub



Private Sub Form_Load()





Set rsNewN = New ADODB.Recordset

'rsNewN.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, KLS_PODR.КОД, Bank.LSCHET, Bank.FIO, Bank.ADR, Bank.KV, Bank.LIFT, Bank.MUSOR, Bank.SELEN, Bank.GVoda, Bank.Otopl, Bank.HVoda, Bank.SSLIV, KLS_PODR.КОД, Bank.Key, Bank.DATA, MainOccupant.Numer, Bank.PERIODOPL FROM (MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) INNER JOIN Bank ON MainOccupant.BanKN = Bank.NewNum"), Mconn, adOpenKeyset, adLockPessimistic
rsNewN.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, KLS_PODR.КОД, Bank.LSCHET, Bank.FIO, Bank.ADR, Bank.KV, Bank.LIFT, Bank.MUSOR, Bank.SELEN, Bank.GVoda, Bank.Otopl, Bank.HVoda, Bank.SSLIV, KLS_PODR.КОД, Bank.Key, Bank.DATA, MainOccupant.Numer, Bank.PERIODOPL, Bank.PLNOM FROM (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) INNER JOIN Bank ON MainOccupant.BanKN = Bank.NewNum"), Mconn
Set fg1.DataSource = rsNewN
rsNewN.Close





MakeWindow Me, True
fg1.Width = Me.Width / 15.40107
fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top

'**************************************************************
'****************************************************************

lblTitle = "Импорт оплаты из банка. Файл > " + BankImport.File1.FileName
Label1.Caption = "Эти строки реестра удалось сопоставить вашим л/сч. Для продолжения нажмите <<Далее>>"


lblTitle.Caption = lblTitle.Caption + "На сумму > " + Str(BankShow.SummI)


End Sub

Private Sub Form_Unload(Cancel As Integer)
'Если импорт тольго по 12 значным номнрам
If MainForm.Bank12 = 1 Then
BankShow12.Show vbModal
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
   Dim i As Long, j As Long, K As Long, rДанные As String
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

Private Sub Image11_Click()

End Sub

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
fg1.Width = Me.Width / 15.40107
   fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
BtnEnh1.Top = Image1.Top

Command3.Left = Image1.Left + Image1.Width
End Sub


