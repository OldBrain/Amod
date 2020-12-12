VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form ArhoPL 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5100
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   8628
   ControlBox      =   0   'False
   Icon            =   "ArhoPL.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   719
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8LCtl.VSFlexGrid VS 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8415
      _cx             =   14843
      _cy             =   6376
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ArhoPL.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
      MergeCompare    =   1
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
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
      Caption         =   "Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
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
      Height          =   168
      Left            =   0
      Picture         =   "ArhoPL.frx":03C9
      Top             =   0
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "И ТОГО:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   4
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   3
      Top             =   4440
      Width           =   1800
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   8370
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   360
      Picture         =   "ArhoPL.frx":0811
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "ArhoPL.frx":0F5B
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "ArhoPL.frx":16A5
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "ArhoPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim s As Double
Dim Er As Label
Dim ARS As ADODB.Recordset
lblTitle.Caption = "История платежей"
MakeWindow Me, True
Set ARS = New ADODB.Recordset
'КоннектА Trim(Str(Year(MainForm.PeriodR))) + "янв.mdb", Filter.Nm

VS.MergeCells = flexMergeRestrictAll

'VS.MergeCellsFixed = flexMergeRestrictAll

VS.Sort = flexSortUseColSort



VS.MergeCol(0) = True
VS.MergeCol(1) = True
VS.MergeCol(2) = True
VS.MergeCol(3) = True

VS.RowHeight(0) = 500
VS.WordWrap = True
VS.Cell(flexcpAlignment, 0, 0, 0, VS.Cols - 1) = flexAlignCenterCenter





End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub

Private Sub VS_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim nr As Long, nc As Long      'при каждом движении мыши вычисляется № строки и столбца
    
    On Error GoTo ex
    Static R As Long, c As Long     'эти №№ изменяются при переходе границы ячейки
    nr = VS.MouseRow:    nc = VS.MouseCol  ' get coordinates
    
    If nr < 1 Or nc = -1 Then
    VS.ToolTipText = ""
    Exit Sub
    End If
    If c <> nc Or R <> nr Then                   ' update tooltip text
        
       If VS.TextMatrix(nr, nc) <> "" Then
        VS.ToolTipText = VS.TextMatrix(nr, nc)
        End If
        R = nr:            c = nc
        DoEvents
    End If
ex:
End Sub
