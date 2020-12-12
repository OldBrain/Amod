VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DropForm2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7170
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4110
   ControlBox      =   0   'False
   DrawWidth       =   2
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      DragIcon        =   "DropF2.frx":0000
      Height          =   375
      Left            =   240
      Picture         =   "DropF2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   3735
   End
   Begin VSFlex8Ctl.VSFlexGrid DG 
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      _cx             =   6588
      _cy             =   10821
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"DropF2.frx":0884
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
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
      AutoSizeMouse   =   0   'False
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
      Height          =   195
      Left            =   0
      Picture         =   "DropF2.frx":09E6
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   3330
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "DropF2.frx":0C30
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   360
      Picture         =   "DropF2.frx":137A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "DropF2.frx":1AC4
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "DropForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim mconn As ADODB.Connection
Dim RSLgota As ADODB.Recordset

Private Sub Command1_Click()
DG_DblClick
End Sub

Private Sub DG_DblClick()

If Arhiv = True Then Exit Sub
'With Me.DG
'DropForm3.DgT.AddItem "След.", 1
'On Error Resume Next
DropForm3.DgT.AddItem "", 0
DropForm3.DgT.TextMatrix(DropForm3.DgT.Row, 0) = Filter.Nm
DropForm3.DgT.Cell(flexcpText, DropForm3.DgT.Row, 1, DropForm3.DgT.Row, 4) = DG.Cell(flexcpText, DG.Row, 1, DG.Row, 2)

DropForm3.DgT.TextMatrix(DropForm3.DgT.Row, 5) = OtheOwner.Othe1

DG.Refresh
'End With

End Sub

Private Sub Form_Load()
MakeWindow Me, True
lblTitle.Caption = "Список льгот"
If Arhiv = True Then Command1.Enabled = False

'Set mconn = New ADODB.Connection
Set RSLgota = New ADODB.Recordset

'conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'conn.Open "data/Kvartplata.mdb"
    
Set RSLgota.ActiveConnection = mConn

RSLgota.CursorType = adOpenForwardOnly
RSLgota.LockType = adLockBatchOptimistic

'RSLgota.Open ("SELECT KLS_PRIV.NAME_KLS FROM KLS_PRIV")
RSLgota.Open ("KLS_PRIV")
Set DG.DataSource = RSLgota


    
    DG.AllowUserResizing = flexResizeBoth
    DG.ExtendLastCol = True
    DG.ExplorerBar = flexExSortShowAndMove
    DG.AutoSearch = flexSearchFromCursor


End Sub



Private Sub Form_Unload(Cancel As Integer)
DropForm3.Hide
End Sub
