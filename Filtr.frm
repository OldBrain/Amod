VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form Filter 
   Caption         =   "Demo - Filtered ADO"
   ClientHeight    =   6285
   ClientLeft      =   2610
   ClientTop       =   2310
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7335
   Begin VSFlex8LCtl.VSFlexGrid FG 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6735
      _cx             =   11880
      _cy             =   8493
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      FormatString    =   ""
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
      ExplorerBar     =   5
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
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Filter"
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
End
Attribute VB_Name = "Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_DS As FlexADO

Private Sub Command1_Click()
    
    ' €сные данные фильтра, восстановлени€
    FG.Cell(flexcpText, 1, 0, 1, FG.Cols - 1) = ""
    FG.FlexDataSource = m_DS
    
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    ' новому фильтру, нужно восстановление
    If Row = 1 Then FG.FlexDataSource = m_DS
    
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    'дл€ этого образца, мы только допускаем редактировать линию фильтра
    If Row <> 1 Then Cancel = True
    
End Sub

Private Sub Form_Load()
'—ортировка//////////////////////////////////
Dim r1 As Long
Dim c1 As Long
Dim Order As Integer

    For r1 = 1 To FG.Rows - 25
        For c1 = 1 To FG.Cols - 1
        FG.TextMatrix(r1, c1) = CInt(Rnd * 9)
        Next c1
    Next r1

'////////////////////////////////////////////////////

FG.AutoSearch = flexSearchFromCursor
FG.ExplorerBar = flexExSortShowAndMove

    ' инициализируйте сетку (дополнительную)
    FG.FixedCols = 0
    FG.Editable = flexEDKbdMouse
    FG.BackColorFrozen = RGB(200, 255, 200)
    
    ' создайте исходный объект заказных данных
    Set m_DS = New FlexADO
    
    ' назначьте этим в сетку
    FG.FlexDataSource = m_DS
    FG.FrozenRows = 1
    
    ' помен€йте размеры базирующийс€ в сначала 6 колонках (дополнительных)
    Dim c%, R%, w%, mw%
    For c = 0 To FG.Cols - 1
        mw = 0
        For R = 0 To 5
            w = TextWidth(FG.TextMatrix(R, c))
            If w > mw Then mw = w
        Next
        FG.ColWidth(c) = mw + 100
    Next
    
End Sub
'////////// —ортировка
Private Sub fg_BeforeSort(ByVal Col As Long, Order As Integer)
    
    ' без флага? приложите заказную сортировку
    If FG.ExplorerBar > &H1000& Then Exit Sub
    
    ' save selection
    Dim R&, c&, rs&, cs&
    FG.GetSelection R, c, rs, cs
    FG.Redraw = flexRDNone
    
    ' приложите сортировку, чтобы не-пустой дипазон
    Dim Row%
    For Row = FG.Rows - 1 To FG.FixedRows Step -1
        If Len(FG.TextMatrix(Row, Col)) Then Exit For
    Next
    If Row > FG.FixedRows Then
        FG.Select FG.FixedRows, Col, Row, Col
        FG.Sort = Order
           
    End If
        
    ' выбор восстановлени€
    FG.Select R, c, rs, cs
    FG.Redraw = flexRDDirect
    
    ' встроенна€ сортировка отмены
    Order = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FG.Move FG.Left, FG.Top, ScaleWidth - FG.Left * 2, ScaleHeight - FG.Left - FG.Top
End Sub
