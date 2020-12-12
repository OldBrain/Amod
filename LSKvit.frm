VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form LSKvit 
   Caption         =   "Отметте лиц.счета для печати"
   ClientHeight    =   7944
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   7944
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   372
      Left            =   480
      TabIndex        =   2
      Top             =   7440
      Width           =   5172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Продолжить"
      Height          =   492
      Left            =   480
      TabIndex        =   1
      Top             =   6840
      Width           =   5172
   End
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   6012
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   5052
      _cx             =   8911
      _cy             =   10604
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
      ExplorerBar     =   0
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
End
Attribute VB_Name = "LSKvit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
rep_kvit.Exit_Me = False
Unload Me
End Sub

Private Sub Command2_Click()

rep_kvit.Exit_Me = True

Unload Me
End Sub

Private Sub Form_Load()

Dim RsLicSh As ADODB.Recordset
'MsgBox (rep_kvit.Label1.Caption)

' Создаем рекордсет лицевых счетов по выбранному дому
Set RsLicSh = New ADODB.Recordset
'Set RsLicSh.ActiveConnection = Mconn
'Set RsLicSh.LockType = adLockBatchOptimistic


RsLicSh.Open ("SELECT MainOccupant.otm, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT From MainOccupant WHERE (((MainOccupant.Dom)=" + rep_kvit.Label1.Caption + "))"), Mconn, , adLockBatchOptimistic



Set VS.DataSource = RsLicSh
End Sub
