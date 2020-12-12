VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Anal_Zatrat 
   Caption         =   "Form3"
   ClientHeight    =   7968
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12060
   LinkTopic       =   "Form3"
   ScaleHeight     =   7968
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   288
      Left            =   6600
      TabIndex        =   11
      Top             =   480
      Width           =   2652
      _ExtentX        =   4678
      _ExtentY        =   508
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin VB.CommandButton Cm3 
      Caption         =   "Печать "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Настройка"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10320
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Расчет"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1572
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   288
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   508
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   288
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   508
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   288
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   508
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   6132
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   11652
      _cx             =   20553
      _cy             =   10816
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      FormatString    =   $"Anal_Zatrat.frx":0000
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
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
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1320
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   -720
         Width           =   312
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Категории затрат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   12
      Top             =   240
      Width           =   2652
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   6480
      X2              =   6480
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Тариф"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5280
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   5160
      X2              =   5160
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   2160
      X2              =   2160
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Категория расчета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Адрес"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   2772
   End
End
Attribute VB_Name = "Anal_Zatrat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Q As String
Public SS As String
Dim RS_zatr As ADODB.Recordset
Dim RS_comboAdr As ADODB.Recordset
Dim RS_comboKat As ADODB.Recordset
Dim RS_comboTar As ADODB.Recordset
Dim RS_comboZ As ADODB.Recordset




Private Sub Cm3_Click()
PrintW.Show
     With PrintW.VP
     
        PrintW.VP.StartDoc
        .FontSize = 12
       .Paragraph = Me.SS + vbNewLine + "_________________________________________________________________"
        .Paragraph = ""
        
        .FontSize = 8
        .RenderControl = fg1.hwnd
        .EndDoc
        
       End With
End Sub

Private Sub Command1_Click()

fg1.Subtotal flexSTSum, 1, 10, fg1.Cols, vbBlue, vbWhite
fg1.Subtotal flexSTSum, 1, 11, fg1.Cols, vbBlue, vbWhite


fg1.Subtotal flexSTSum, 1, 12, fg1.Cols, vbBlue, vbWhite
fg1.Subtotal flexSTSum, 0, 12, fg1.Cols, vbWhite, vbBlue, True, "ВСЕГО", , True

fg1.Subtotal flexSTSum, 1, 13, fg1.Cols, vbBlue, vbWhite
fg1.Subtotal flexSTSum, 0, 13, fg1.Cols, vbWhite, vbBlue, True, "ВСЕГО", , True

fg1.Subtotal flexSTSum, 1, 14, fg1.Cols, vbBlue, vbWhite
fg1.Subtotal flexSTSum, 0, 14, fg1.Cols, vbWhite, vbBlue, True, "ВСЕГО", , True

fg1.ColWidth(0) = 700

'FG1.Subtotal flexSTSum, 1, 11, FG1.Cols, vbBlue, vbWhite, False, "И того"
'FG1.Subtotal flexSTSum, 1, 12, FG1.Cols, vbBlue, vbWhite, False, "И того"
'FG1.Subtotal flexSTSum, 1, 13, FG1.Cols, vbBlue, vbWhite, False, "И того"
'FG1.Subtotal flexSTSum, 1, 14, FG1.Cols, vbBlue, vbWhite, False, "И того"


'Dim i As Integer
'Ur1 = 10
'Me.Об 1
'zapfg (Q)
End Sub

Private Sub Command2_Click()
Dim i As Integer
Ur1 = 10
Me.Об 1
End Sub

Private Sub DataCombo1_Validate(Cancel As Boolean)
' MsgBox (Me.DataCombo1.BoundText)

zapfg (Q)
End Sub



Private Sub DataCombo2_Validate(Cancel As Boolean)
zapfg (Q)
End Sub

Private Sub DataCombo3_Validate(Cancel As Boolean)
zapfg (Q)

End Sub

Private Sub DataCombo4_Validate(Cancel As Boolean)
zapfg (Q)
End Sub

Private Sub Form_Load()
Me.Caption = "Фильтр>"
Set RS_zatr = New ADODB.Recordset
Set RS_comboAdr = New ADODB.Recordset
Set RS_comboKat = New ADODB.Recordset
Set RS_comboTar = New ADODB.Recordset
Set RS_comboZ = New ADODB.Recordset

RS_comboAdr.Open ("SELECT kls_podr.NAIM_KLS, kls_podr.КОД From kls_podr"), Mconn, adOpenStatic, adLockBatchOptimistic
RS_comboKat.Open ("SELECT Kategor.Код, Kategor.Name_Kategor FROM Kategor"), Mconn, adOpenStatic, adLockBatchOptimistic
RS_comboTar.Open ("SELECT Tarif.Value From Tarif GROUP BY Tarif.Value"), Mconn, adOpenStatic, adLockBatchOptimistic
RS_comboZ.Open ("SELECT Schet.Schet, Schet.Schet_Name From Schet GROUP BY Schet.Schet, Schet.Schet_Name"), Mconn, adOpenStatic, adLockBatchOptimistic


' Заполняем комбо ЗАТРАТ
With DataCombo4

        Set .DataSource = RS_comboZ
        Set .RowSource = RS_comboZ
        .DataField = "Schet"
        .ListField = "Schet_Name"
        
        .BoundColumn = "Schet"
                    
    End With


' Заполняем комбо Адрес
With DataCombo1

        Set .DataSource = RS_comboAdr
        Set .RowSource = RS_comboAdr
        .DataField = "КОД"
        .ListField = "NAIM_KLS"
        
        .BoundColumn = "КОД"
                    
    End With

' Заполняем комбо категория расчета
With DataCombo2
        Set .DataSource = RS_comboKat
        Set .RowSource = RS_comboKat
        .DataField = "КОД"
        .ListField = "Name_Kategor"
        
        .BoundColumn = "КОД"

End With

' Заполняем комбо тарифы
With DataCombo3
        Set .DataSource = RS_comboTar
        Set .RowSource = RS_comboTar
        .DataField = "Value"
        .ListField = "Value"
        
        .BoundColumn = "Value"

End With

'RS_combo.Close

'Set Me.DataCombo1.DataSource = RS_combo



'Filt = "Where (([arh_rep]![Tip] Like '*'))"
Filt = " WHERE (((Adding.SummaI)<>0)"
'Filt = Filt + " and Adding.Tip Like '+')"
Filt = Filt + ")"
' Резерв Q = "SELECT MainOccupant.Dom, kls_podr.NAIM_KLS, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Tip as 'Начислено(+)/оплачено(-)', zatr_tarif.Schet, zatr_tarif.Schet_Name, Adding.SummaI as 'Начислено/оплачено',zatr_tarif.Summa, zatr_tarif.Procent as 'Процентная  доля затрат', IIf([adding.Tip]='+',[SummaI]*[Procent]/100,0) AS Zatrat_N, IIf([adding.Tip]='-',[SummaI]*[Procent]/100,0) AS Zatrat_O, IIf([adding.Tip]='S',[SummaI]*[Procent]/100,0) AS Zatrat_S, Adding.TipKvKod, Adding.TipDomKod, Adding.KodN, Adding.KodKv, Adding.FLOOR, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.kod_tar FROM kls_podr INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN zatr_tarif ON Adding.kod_tar = zatr_tarif.kod_tar) ON MainOccupant.Numer = Adding.KodKv) ON kls_podr.КОД = MainOccupant.Dom " + Filt
Q = "SELECT MainOccupant.Dom, kls_podr.NAIM_KLS, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Tip as 'Начислено(+)/оплачено(-)', zatr_tarif.Schet, zatr_tarif.Schet_Name, Adding.SummaI as 'ВСЕГО Нач/опл', zatr_tarif.Summa as 'Доля затр (РУБ)', zatr_tarif.Procent as 'Доля затр (%)', IIf([adding.Tip]='+',[SummaI]*[Procent]/100,0) AS 'Доля затр ПЛАН(начислено)', IIf([adding.Tip]='-',[SummaI]*[Procent]/100,0) AS 'Доля затр ФАКТ(оплачено)', IIf([adding.Tip]='S',[SummaI]*[Procent]/100,0) AS 'Доля затр (субсидии)', Adding.TipKvKod, Adding.TipDomKod, Adding.KodN, Adding.KodKv, Adding.FLOOR, Adding.Shc_old, Adding.Shc_new, Adding.Sch , Adding.kod_tar FROM kls_podr INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN zatr_tarif ON Adding.kod_tar = zatr_tarif.kod_tar) ON MainOccupant.Numer = Adding.KodKv) ON kls_podr.КОД = MainOccupant.Dom " + Filt

', zatr_tarif.Summa as 'Доля затр (РУБ)', zatr_tarif.Procent as 'Доля затр (%)', IIf([adding.Tip]='+',[SummaI]*[Procent]/100,0) AS 'Доля затр ПЛАН(начислено)', IIf([adding.Tip]='-',[SummaI]*[Procent]/100,0) AS 'Доля затр ФАКТ(оплачено)', IIf([adding.Tip]='S',[SummaI]*[Procent]/100,0) AS 'Доля затр (субсидии)'
fg1.Sort = flexSortGenericAscending

RS_zatr.Open (Q), Mconn
Set fg1.DataSource = RS_zatr

' Устанавливаем возможность регулировать ширину и толщ. колонок при помощи мыши
fg1.AllowUserResizing = flexResizeBoth


'Уст возможность тоскать колонки
fg1.ExplorerBar = flexExMove

' Устанавливаем атрибут ПЕРЕНОСИТЬ ПО СЛОВАМ при  регулировке ширины и толщ. колонок при помощи мыши
fg1.WordWrap = True

'Убираем колонку Адрес и код категории
fg1.ColHidden(1) = True
fg1.ColHidden(3) = True

'Убираем колонку тариф
fg1.ColHidden(5) = True
'Убираем колонку +/-
fg1.ColHidden(6) = True
'Убираем колонку  код тарифа
fg1.ColHidden(23) = True

'Убираем колонку  код затрат
fg1.ColHidden(7) = True

' Прячем прочие ненужные колонки
fg1.ColHidden(15) = True
fg1.ColHidden(16) = True
fg1.ColHidden(17) = True
fg1.ColHidden(18) = True
fg1.ColHidden(19) = True
fg1.ColHidden(20) = True
fg1.ColHidden(21) = True
fg1.ColHidden(22) = True



'Шилина колонок
fg1.ColWidth(10) = 600
fg1.ColWidth(11) = 600
fg1.ColWidth(12) = 900
fg1.ColWidth(13) = 1000
fg1.ColWidth(14) = 1000
' ВЫСОТА ВЕРХНЕЙ СТРОКИ
fg1.RowHeight(0) = 700
End Sub
Private Sub zapfg(Q1 As String)

Filt = " WHERE ((Adding.SummaI<>0)"

'Фильтруем Адрес
If Me.DataCombo1.BoundText <> "0" Then
Filt = Filt + " and Cstr(MainOccupant.Dom) Like '" + Me.DataCombo1.BoundText + "'"
fg1.ColHidden(2) = True
Else
Filt = Filt + " and Cstr(MainOccupant.Dom) Like '%'"
fg1.ColHidden(2) = False
End If

'Фильтруем Категорию расчета
If Me.DataCombo2.BoundText <> "0" Then
Filt = Filt + " and Cstr(Adding.KodKat) Like '" + Me.DataCombo2.BoundText + "'"
Else
Filt = Filt + " and Cstr(Adding.KodKat) Like '%'"
End If


'Фильтруем Тариф
If Me.DataCombo3.BoundText <> "0" Then
Filt = Filt + " and Cstr(Adding.tarif) Like '" + Me.DataCombo3.BoundText + "'"
'Убираем колонку тариф
fg1.ColHidden(5) = True
Else
Filt = Filt + " and Cstr(Adding.tarif) Like '%'"
End If



'Фильтруем Категорию ЗАТРАТ
If Me.DataCombo4.BoundText <> "0" Then
Filt = Filt + " and Cstr(zatr_tarif.Schet) Like '" + Me.DataCombo4.BoundText + "'"

End If

'Добавляем последнюю скобку в фильтр
Filt = Filt + ")"


'MsgBox (Filt)

Q = "SELECT MainOccupant.Dom, kls_podr.NAIM_KLS, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Tip as 'Начислено(+)/оплачено(-)', zatr_tarif.Schet, zatr_tarif.Schet_Name, Adding.SummaI as 'ВСЕГО Нач/опл', zatr_tarif.Summa as 'Доля затр (РУБ)', zatr_tarif.Procent as 'Доля затр (%)', IIf([adding.Tip]='+',[SummaI]*[Procent]/100,0) AS 'Доля затр ПЛАН(начислено)', IIf([adding.Tip]='-',[SummaI]*[Procent]/100,0) AS 'Доля затр ФАКТ(оплачено)', IIf([adding.Tip]='S',[SummaI]*[Procent]/100,0) AS 'Доля затр (субсидии)', Adding.TipKvKod, Adding.TipDomKod, Adding.KodN, Adding.KodKv, Adding.FLOOR, Adding.Shc_old, Adding.Shc_new, Adding.Sch , Adding.kod_tar FROM kls_podr INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN zatr_tarif ON Adding.kod_tar = zatr_tarif.kod_tar) ON MainOccupant.Numer = Adding.KodKv) ON kls_podr.КОД = MainOccupant.Dom " + Filt
Q1 = Q





RS_zatr.Close
Set fg1.DataSource = Nothing



SS = "Адрес=" + Me.DataCombo1.Text + " /Категория=" + Me.DataCombo2.Text + " /Тариф=" + Me.DataCombo3.Text + " /Затраты =" + Me.DataCombo4.Text
MsgBox ("Устанавливаем фильтр > " + SS)

RS_zatr.Open (Q1), Mconn
Set fg1.DataSource = RS_zatr


'FG1.Refresh

'Убираем колонку Адрес
fg1.ColHidden(1) = True
If Me.DataCombo1.BoundText <> "0" Then
fg1.ColHidden(2) = True
Else
fg1.ColHidden(2) = False
End If

'Убираем колонку категория расчета
fg1.ColHidden(3) = True
If Me.DataCombo2.BoundText <> "0" Then

fg1.ColHidden(4) = True
Else
fg1.ColHidden(4) = False
End If


'Убираем колонку тариф
If Me.DataCombo3.BoundText <> "0" Then
fg1.ColHidden(5) = True
Else
fg1.ColHidden(5) = False
End If

'Убираем колонку код тарифа
fg1.ColHidden(23) = True



'Убираем колонку +/-
fg1.ColHidden(6) = True

'Убираем колонку  код затрат
fg1.ColHidden(7) = True

' Прячем прочие ненужные колонки
fg1.ColHidden(15) = True
fg1.ColHidden(16) = True
fg1.ColHidden(17) = True
fg1.ColHidden(18) = True
fg1.ColHidden(19) = True
fg1.ColHidden(20) = True
fg1.ColHidden(21) = True
fg1.ColHidden(22) = True

'Ширина колонок
fg1.ColWidth(10) = 600
fg1.ColWidth(11) = 600
fg1.ColWidth(12) = 900
fg1.ColWidth(13) = 1000
fg1.ColWidth(14) = 1000


Me.Caption = SS
'Me.Caption = Me.Caption + " /Категория=" + Me.DataCombo2.Text



'Группировка
 fg1.MergeCells = flexMergeRestrictAll
 fg1.MergeCol(-1) = True
 fg1.MergeCol(fg1.Cols - 1) = False
        ' установите слияние ячейки (все колонны)
        fg1.MergeCells = flexMergeRestrictAll
        fg1.MergeCol(-1) = True


'FG1.MergeCells = flexMergeRestrictAll
'FG1.MergeCol(-1) = True
' FG1.Refresh

'FG1.Sort = flexSortGenericAscending

End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu_zatr.Show
End Sub

Public Sub Об(ur As Integer)


On Error GoTo er1
'MsgBox (Str(ur))


         '*************расчет I-ой колонки ***************************
         
     If ur = 0 Then Exit Sub
     If ur = 1 Then
     
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear

     
        For i = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
         Next
          ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
     End If
        
     If ur = 2 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     
     
     For i = 2 To fg1.Cols - 1
        'fg1.Subtotal flexSTSum, -1, i, , RGB(200, 255, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
        fg1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, True
        
    Next
     ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
        
        
     If ur = 3 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
       
        
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     For i = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, False, "И того:"
        fg1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, False
        fg1.Subtotal flexSTSum, 2, i, , RGB(200, 250, 200), vbBlack, False
       Next
        ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If

 If ur = 4 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     
     
     
     For i = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
        fg1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, False
        fg1.Subtotal flexSTSum, 2, i, , RGB(200, 250, 200), vbBlack, False
        fg1.Subtotal flexSTSum, 3, i, , RGB(100, 200, 200), vbBlack, False
           Next
        ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If

If ur = 5 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     
     
     
     For i = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
        fg1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, False
        fg1.Subtotal flexSTSum, 2, i, , RGB(200, 250, 200), vbBlack, False
        fg1.Subtotal flexSTSum, 3, i, , RGB(100, 200, 200), vbBlack, False
        fg1.Subtotal flexSTSum, 4, i, , RGB(300, 100, 200), vbBlack, False
       Next
        ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If
     
     For i = 500 To 1000
    Pod.ProgressBar1.Value = i
    For j = 1 To 100000
    Next j
   Next

     
     Exit Sub
er1:
If Err.Number = 381 Then
MsgBox "Нет данных для отчета с выбранными параметрами " + Me.Caption
Unload Pod
Unload Jdite
Unload Uroven
MainMenu.Enabled = True
MainMenu.Show
Unload Me

Exit Sub
Else
MsgBox Err.Description
End If
End Sub

