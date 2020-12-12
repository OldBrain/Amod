VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form VibDom 
   ClientHeight    =   7272
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5808
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "VibDom.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   2  'CenterScreen
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Width           =   5535
      _ExtentX        =   9758
      _ExtentY        =   656
      Caption         =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      _cx             =   9763
      _cy             =   10821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      FormatString    =   $"VibDom.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   0
      Picture         =   "VibDom.frx":03E0
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
      Picture         =   "VibDom.frx":0B2A
      Top             =   0
      Width           =   156
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Выбор домов для отчета"
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   5730
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   4800
      Picture         =   "VibDom.frx":0D74
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   2880
      Picture         =   "VibDom.frx":14BE
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "VibDom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nabor(100) As Integer
Dim Sqf As String

Private Sub Form_Load()
MakeWindow Me, False
Dim rsDom As ADODB.Recordset
Set rsDom = New ADODB.Recordset

rsDom.Open ("SELECT KLS_PODR.КОД AS Код, KLS_PODR.NAIM_KLS AS Улица, KLS_PODR.Num AS [№] FROM KLS_PODR Order by KLS_PODR.NAIM_KLS"), Mconn
Set VS.DataSource = rsDom


VS.Editable = flexEDKbdMouse


VS.Cell(flexcpChecked, 1, 0, VS.Rows - 1, 0) = flexUnchecked



End Sub

Private Sub VS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If Col <> 0 Then Cancel = True

End Sub

Private Sub VS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If Col <> 0 Then Cancel = True
End Sub

Private Sub xpcmdbutton1_Click()
ic = 0

Sqf = "HAVING ((("
Analizlgot.Titl = "Свод по домам:"
For rw = 1 To VS.Rows - 1

If VS.Cell(flexcpChecked, rw, 0) = flexChecked Then
ic = ic + 1
Analizlgot.Titl = Analizlgot.Titl + VS.TextMatrix(rw, 2) + ", "
If ic > 1 Then Sqf = Sqf + " OR (("

Sqf = Sqf + "KLS_PODR.КОД)=" + VS.TextMatrix(rw, 1) + ")"
End If
Next

Analizlgot.Titl = Analizlgot.Titl + " за " + MainMenu.Command13.Caption

If iс > 1 Then Sqf = Sqf + ")))" Else Sqf = Sqf + ")"


'MsgBox Sqf


sq = "SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.BanKN AS N, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.kv_num AS [Кв №], MainOccupant.COMSPACE AS [Общая пл], MainOccupant.NLODGERF AS Прописано, Sum((Adding!SaldoN*1000/Adding!Kol)/1000) AS [Саольдо нач], Sum(IIf(Adding!Tip='+',[SummaI],0)) AS Начислено, Sum(IIf(Adding!Tip='s',[SummaI],0)) AS Субсидии, Sum(IIf(Adding!Tip='-',[SummaI],0)) AS Оплата, Sum((Adding!SaldoK*1000/Adding!Kol)/1000) AS [Саольдо кон], MainOccupant.Dog AS [Абон книжка] FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД GROUP BY KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.BanKN, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.COMSPACE, MainOccupant.NLODGERF, MainOccupant.Dog " + Sqf


'************
Analizlgot.G = 17
Analizlgot.StrSQL = sq
Analizlgot.Show

Analizlgot.fg1.ColHidden(1) = True
Analizlgot.fg1.ColHidden(2) = True
Analizlgot.fg1.ColHidden(3) = True

Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"
Analizlgot.fg1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True
If MainForm.Dog = 1 Then Analizlgot.fg1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
'*********


'SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.BanKN AS N, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.kv_num AS [Кв №], MainOccupant.COMSPACE AS [Общая пл], MainOccupant.NLODGERF AS Прописано, Sum((Adding!SaldoN*1000/Adding!Kol)/1000) AS [Саольдо нач], Sum(IIf(Adding!Tip='+',[SummaI],0)) AS Начислено, Sum(IIf(Adding!Tip='s',[SummaI],0)) AS Субсидии, Sum(IIf(Adding!Tip='-',[SummaI],0)) AS Оплата, Sum((Adding!SaldoK*1000/Adding!Kol)/1000) AS [Саольдо кон], MainOccupant.Dog AS [Абон книжка]
'FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД
'GROUP BY KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.BanKN, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.COMSPACE, MainOccupant.NLODGERF, MainOccupant.Dog
'HAVING (((KLS_PODR.КОД)=1)) OR (((KLS_PODR.КОД)=2)) OR (((KLS_PODR.КОД)=3));


Unload Form8

Unload Me
End Sub
