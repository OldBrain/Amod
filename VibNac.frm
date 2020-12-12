VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form VibNac 
   Caption         =   "Form5"
   ClientHeight    =   4428
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5064
   LinkTopic       =   "Form5"
   ScaleHeight     =   4428
   ScaleWidth      =   5064
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Готово"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   4932
   End
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   2532
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4692
      _cx             =   8276
      _cy             =   4466
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
      FormatString    =   $"VibNac.frx":0000
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
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Не более 5"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Пожалуйста, отметте платежи для выгрузки"
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
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4932
   End
End
Attribute VB_Name = "VibNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()



'VS.Refresh

'Dim Nabor(10, 1, 0) As String
'ReDim Nabor(10, 1, 0)

'Выбираем открыженные
For rw = 1 To VS.Rows - 1

'MsgBox (VS.TextMatrix(rw, 0))

'Проверяем на чек
    If VS.TextMatrix(rw, 0) = -1 Then
        Nabor(rw) = VS.TextMatrix(rw, 1)
        
        'MsgBox (Nabor(rw))
    End If
Next rw


Unload Me

'MsgBox (rw)
End Sub

Private Sub Form_Load()

Dim Sqf As String
'MakeWindow Me, False
Dim rsDom As ADODB.Recordset
Set rsDom = New ADODB.Recordset

rsDom.Open ("SELECT Adding.KodN, Adding.NameN, Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI] From Adding GROUP BY Adding.KodN, Adding.NameN, Adding.Tip HAVING (((Adding.Tip)='-') AND ((Sum(Adding.SummaI))>0))"), Mconn
'rsDom.Open ("SELECT KLS_PODR.КОД AS Код, KLS_PODR.NAIM_KLS AS Улица, KLS_PODR.Num AS [№] FROM KLS_PODR Order by KLS_PODR.NAIM_KLS"), Mconn


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


