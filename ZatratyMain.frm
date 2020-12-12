VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form ZatratyMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8436
   ClientLeft      =   0
   ClientTop       =   300
   ClientWidth     =   6828
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ZatratyMain.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   703
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid Fg1 
      Height          =   7692
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6612
      _cx             =   11663
      _cy             =   13568
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
      ForeColor       =   -2147483646
      BackColorFixed  =   -2147483648
      ForeColorFixed  =   -2147483645
      BackColorSel    =   -2147483630
      ForeColorSel    =   255
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   16777215
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
      FormatString    =   $"ZatratyMain.frx":0ECA
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
      Height          =   192
      Left            =   5760
      Picture         =   "ZatratyMain.frx":0FBF
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   192
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "АРМ ""Квартплата + "" Отчеты и Анализ"
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
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   6216
   End
   Begin VB.Image imgTitleMain 
      Height          =   360
      Left            =   3480
      Picture         =   "ZatratyMain.frx":1501
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   288
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   360
      Picture         =   "ZatratyMain.frx":1C4B
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   6120
      Picture         =   "ZatratyMain.frx":2395
      Top             =   0
      Width           =   228
   End
   Begin VB.Menu Plan 
      Caption         =   "План"
      Index           =   1
   End
   Begin VB.Menu Schet 
      Caption         =   "Справочник затрат"
      Index           =   0
   End
End
Attribute VB_Name = "ZatratyMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_kat As ADODB.Recordset

Private Sub FG1_Click()
If fg1.Row <> 0 Then ZatrPopUp.lblTitle.Caption = "Коментарий по адресу " + fg1.TextMatrix(fg1.Row, 3) Else MsgBox fg1.Row

End Sub

Private Sub Fg1_DblClick()
ZatrAdding.Show
End Sub

Private Sub Fg1_GotFocus()

' Распологаем окно коментария относительно главного окна
ZatrPopUp.Show
ZatrPopUp.Height = Me.Top
ZatrPopUp.Width = Me.Width + Me.Left

ZatrPopUp.Enabled = False
ZatrPopUp.Refresh
MakeWindow ZatrPopUp, True


End Sub

Private Sub Form_Activate()
'ZatrPopUp.ScaleHeight = ZatratyMain.Top
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
      Unload Me
      MainMenu.Show
   End If

End Sub
Private Sub Form_Load()
'КоннектЗ

MakeWindow Me, True
lblTitle.Caption = "Расчет затрат"

Me.KeyPreview = True
 
 
 fg1.Editable = False
' open connection
  ' Set mconn = New ADODB.Connection
'  mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
    
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
 
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
rs_kat.Open ("SELECT KLS_PODR.КОД, KLS_PODR.SaldoN, KLS_PODR.NAIM_KLS,  KLS_PODR.SaldoK  From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")
Set fg1.DataSource = rs_kat



Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
Combo_RS.Open "TipDom"



' правопреемник recordset в сетку
   
   fg1.FocusRect = 3
    'flexFocusSolid
    fg1.Editable = True
    fg1.DataMode = flexDMBound
        
    fg1.AutoSearch = flexSearchFromCursor
    fg1.ExplorerBar = flexExSortShowAndMove
    
    
    ' Cвойства, свойства необходимые для сортировки в этом гриде не работают
    ' из за строки поиска
    fg1.AllowUserResizing = flexResizeBoth
    fg1.ExtendLastCol = True
    
    

    
    
    





'If Me.Width > Application.UsableWidth Then Me.Width = Application.UsableWidth - 10

End Sub


Private Sub imgTitleHelp_Click()

MainMenu.Show
Unload Me
End Sub

Private Sub Plan_Click(Index As Integer)

ZPlan.Dom = Me.fg1.TextMatrix(fg1.Row, 1)

ZPlan.Show 1, Me
End Sub

Private Sub Schet_Click(Index As Integer)
Schet1.Show 1, Me
End Sub
