VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form11 
   Caption         =   "����� �� ������� <���������� �������>"
   ClientHeight    =   7230
   ClientLeft      =   3255
   ClientTop       =   2475
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleMode       =   0  '����������������
   ScaleWidth      =   13821.99
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "������ �����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����������� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������ �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������������ ��������� ��������� �����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Bindings        =   "Form_grupp.frx":0000
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   11775
      _cx             =   20770
      _cy             =   8705
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Form_grupp.frx":001F
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
      MergeCompare    =   1
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
      ShowComboButton =   0
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   "VBA"
      ComboSearch     =   1
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
   Begin VB.Frame Frame1 
      Caption         =   "���� ������������ ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '���������
      Caption         =   $"Form_grupp.frx":0105
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      TabIndex        =   9
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
Private Sub Command1_Click()
' initialize the control
fg.MergeCells = 0

       fg.Cols = 5
       fg.FixedCols = 0
       fg.GridLinesFixed = flexGridExplorer
       fg.AllowUserResizing = flexResizeBoth
       fg.ExplorerBar = flexExMove
      'fg.Editable = 2
       fg.Redraw = True
       fg.MergeCol(-1) = True
       fg.Subtotal flexSTClear
       fg.OutlineBar = 0
       Form11.Hide
       Form11.Show
       End Sub

Private Sub Command2_Click()
Dim i As Integer

' suspend repainting to get more speed
        fg.Redraw = False
        fg.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg.Select 1, 0, 1, fg.Cols - 1
        fg.Sort = flexSortGenericAscending
        fg.Select 1, 0
        ' calculate subtotals
        fg.Subtotal flexSTClear
        
For i = 1 To fg.Cols - 1
         '*************������ I-�� ������� ***************************
        fg.Subtotal flexSTSum, -1, i, , 1, vbWhite, True
        fg.Subtotal flexSTSum, 0, i, , vbRed, vbWhite, True
        fg.Subtotal flexSTSum, 1, i, , vbBlue, vbWhite, True
        'MsgBox i
    Next i
        
        ' autosize
        fg.AutoSize 0, fg.Cols - 1, , 300
        ' turn repainting back on
        fg.OutlineBar = flexOutlineBarComplete
        fg.Redraw = True
End Sub

Private Sub Command3_Click()
fg.OutlineBar = 0
fg.FixedCols = 0
       fg.GridLinesFixed = flexGridExplorer
       fg.AllowUserResizing = flexResizeBoth
       fg.ExplorerBar = flexExMove
      'fg.Editable = 2
       fg.Redraw = True
       fg.MergeCol(-1) = True
       fg.Subtotal flexSTClear
       fg.OutlineBar = 0
       Form11.Refresh
fg.Refresh
End Sub

Private Sub Command4_Click()
fg.MergeCells = flexMergeRestrictAll
 fg.MergeCol(-1) = True
 fg.Refresh
 
 End Sub

Private Sub Command5_Click()
fg.MergeCells = 0
fg.MergeCol(-1) = True
fg.OutlineBar = 0
fg.FixedCols = 0
       fg.GridLinesFixed = flexGridExplorer
       fg.AllowUserResizing = flexResizeBoth
       fg.ExplorerBar = flexExMove
      'fg.Editable = 2
       fg.Redraw = True
       fg.MergeCol(-1) = True
       fg.Subtotal flexSTClear
       fg.OutlineBar = 0
       Form11.Refresh

fg.Refresh
 
End Sub

Private Sub Command6_Click()
Form11.Hide
Form1.Show
End Sub

 Private Sub Form_Load()

    
        ' initialize the control
        fg.Cols = 5
        fg.FixedCols = 0
        fg.GridLinesFixed = flexGridExplorer
        fg.AllowUserResizing = flexResizeBoth
        fg.ExplorerBar = flexExMove
        'fg.Editable = 2

 fg.MergeCells = flexMergeRestrictAll
 fg.MergeCol(-1) = True
 fg.MergeCol(fg.Cols - 1) = False




        ' ���������� ������� ������ (��� �������)
        'fg.MergeCells = flexMergeRestrictAll
        fg.MergeCol(-1) = True
 
  
     
    End Sub






Private Sub fg_AfterMoveColumn(ByVal Col As Long, Position As Long)

       ' sort the data from first to last column
     fg.Select 1, 0, 1, fg.Cols - 1
     fg.Sort = flexSortGenericAscending
     fg.Select 1, 0

'***********************************************************
  
'*************************************************************
 End Sub

' Private Sub fg_BeforeMoveColumn(ByVal Col As Long, Position As Long)
        ' don't move sales figures
 '       If Col = fg.Cols - 1 Then Position = -1
  '  End Sub


