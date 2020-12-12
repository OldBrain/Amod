VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form111 
   ClientHeight    =   7230
   ClientLeft      =   3255
   ClientTop       =   2475
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleMode       =   0  'Пользовательское
   ScaleWidth      =   13821.99
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Убрать правую колонку"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Печать"
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
      Caption         =   "Закрыть"
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
      Caption         =   "Отмена группировки"
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
      Caption         =   "Группировка данных"
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
      Caption         =   "Отмена расчета"
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
      Caption         =   "Расчет"
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
      Caption         =   "Восстановить начальные парамнтры формы"
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
      Height          =   4935
      Left            =   120
      Negotiate       =   -1  'True
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
      FormatString    =   $"Form_grupp1.frx":0000
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
      DataMode        =   2
      VirtualData     =   -1  'True
      DataMember      =   ""
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
      Caption         =   "Блок формирования отчета"
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
      Alignment       =   2  'Центровка
      Caption         =   $"Form_grupp1.frx":00E6
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
Attribute VB_Name = "Form111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sq As String




 
Private Sub Command1_Click()
' initialize the control
FG.MergeCells = 0
       FG.Cols = 10
       FG.FixedCols = 0
       FG.GridLinesFixed = flexGridExplorer
       FG.AllowUserResizing = flexResizeBoth
       FG.ExplorerBar = flexExMove
      'fg.Editable = 2
       FG.Redraw = True
       FG.MergeCol(-1) = True
       FG.Subtotal flexSTClear
       FG.OutlineBar = 0
        '********* сортировка ******************
       
       FG.Sort = flexSortGenericAscending
       FG.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
       
       Form111.Hide
       Form111.Show
       End Sub

Private Sub Command2_Click()
Dim i As Integer

' suspend repainting to get more speed
        FG.Redraw = False
        FG.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        FG.Select 1, 0, 1, FG.Cols - 1
        FG.Sort = flexSortGenericAscending
        FG.Select 1, 0
        ' calculate subtotals
        FG.Subtotal flexSTClear
        
'fg.ColFormat(i) = "$(#,###.00)" ' set format for calculated totals
FG.Subtotal flexSTClear         ' remove old values
        
'For i = 1 To fg.Cols - 1

'fg.RowOutlineLevel=
'fg.SubtotalPosition = 1
'fg.Subtotal flexSTSum
         '*************расчет I-ой колонки ***************************
        FG.Subtotal flexSTSum, -1, 2, , vbCyan, vbBlue, True, , , True
        FG.Subtotal flexSTSum, 0, 2, , vbMagenta, vbWhite, True, , , True
        FG.Subtotal flexSTSum, 1, 2, , vbBlue, vbWhite, True, , , True
        'MsgBox i
        FG.AutoSize 0, FG.Cols - 1, , 300
        
        
        
        'fg.Subtotal flexSTSum, -1, i, , 1, vbWhite, True
        'fg.Subtotal flexSTSum, 0, i, , vbRed, vbWhite, True
        'fg.Subtotal flexSTSum, 1, i, , vbBlue, vbWhite, True
        
    'Next i
        
        ' autosize
        FG.AutoSize 0, FG.Cols - 1, , 300
        ' turn repainting back on
        FG.OutlineBar = flexOutlineBarComplete
        FG.Redraw = True
        
        
        
End Sub

Private Sub Command3_Click()
FG.OutlineBar = 0
FG.FixedCols = 0
       FG.GridLinesFixed = flexGridExplorer
       FG.AllowUserResizing = flexResizeBoth
       FG.ExplorerBar = flexExMove
      'fg.Editable = 2
       FG.Redraw = True
       FG.MergeCol(-1) = True
       FG.Subtotal flexSTClear
       FG.OutlineBar = 0
       FG.Subtotal flexSTClear
       Form111.Refresh
FG.Refresh
End Sub

Private Sub Command4_Click()
FG.MergeCells = flexMergeRestrictAll
 FG.MergeCol(-1) = True
 FG.Refresh
 
 End Sub

Private Sub Command5_Click()
FG.MergeCells = 0
FG.MergeCol(-1) = True
FG.OutlineBar = 0
FG.FixedCols = 0
       FG.GridLinesFixed = flexGridExplorer
       FG.AllowUserResizing = flexResizeBoth
       FG.ExplorerBar = flexExMove
      'fg.Editable = 2
       FG.Redraw = True
       FG.MergeCol(-1) = True
       FG.Subtotal flexSTClear
       FG.OutlineBar = 0
       Form111.Refresh

FG.Refresh
 
End Sub

Private Sub Command6_Click()
Form111.Hide
'AcnRst = ""
'Set fg.DataSource = AcnRst
Form1.Show

End Sub

Private Sub Command7_Click()
PrintW.Show

End Sub

Private Sub Command8_Click()
FG.Cols = FG.Cols - 1

End Sub

 Private Sub Form_Load()
    
    '////////////////////Блок Данные////////////////////////////
Dim Acn As ADODB.Connection
Set Acn = New ADODB.Connection
'Acn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data\KV.mdb;Persist Security Info=True"
Acn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data\KV.mdb;Persist Security Info=True"
'Acn.Mode = adModeWrite
 

Acn.Open
Dim AcnRst As ADODB.Recordset
Set AcnRst = New ADODB.Recordset
Dim sq As String

Call MakeSQ(sq)
Form6.Hide
Set AcnRst.ActiveConnection = Acn
AcnRst.Open (sq)
' activate merging for all columns
 FG.MergeCells = flexMergeRestrictAll
 FG.MergeCol(-1) = True
' ******** Установка Data Sourse *************
 FG.VirtualData = True
 Set FG.DataSource = AcnRst
  FG.DataMode = flexDMBound
 '********************************************
'******** Закрываем RecordSet *************
AcnRst.Close
Set AcnRst = Nothing
'******** Закрываем Connect *************
Acn.Close
Set Acn = Nothing
'/////////////Конец блока данных ///////////////////////////////////////
        ' initialize the control
        FG.Cols = 10
        FG.FixedCols = 0
        FG.GridLinesFixed = flexGridExplorer
        FG.AllowUserResizing = flexResizeBoth
        FG.ExplorerBar = flexExMove
        'fg.Editable = 2

        FG.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
 FG.MergeCells = flexMergeRestrictAll
 FG.MergeCol(-1) = True
 FG.MergeCol(FG.Cols - 1) = False
 ' установите слияние ячейки (все колонны)
 'fg.MergeCells = flexMergeRestrictAll
  FG.MergeCol(-1) = True
    End Sub




 

Private Sub fg_AfterMoveColumn(ByVal Col As Long, Position As Long)




       ' sort the data from first to last column
     FG.Select 1, 0, 1, FG.Cols - 1
     FG.Sort = flexSortGenericAscending
     FG.Select 1, 0

'***********************************************************
  
'*************************************************************
 End Sub

' Private Sub fg_BeforeMoveColumn(ByVal Col As Long, Position As Long)
        ' don't move sales figures
 '       If Col = fg.Cols - 1 Then Position = -1
  '  End Sub

Sub MakeSQ(sq)
Dim Tbl As String

Tbl = "J_ALL"
sq = "SELECT "

' ************ 1
If Form6.Check1.Value Then
            sq = sq + Tbl + ".[Наименование_льготы], "
        Else
            sq = sq
        End If
' ************ 2
If Form6.Check2.Value Then
            sq = sq + Tbl + ".NAIM_KLS, "
        Else
            sq = sq
        End If
        
' ************ 3
If Form6.Check3.Value Then
            sq = sq + Tbl + ".[ФИО], "
        Else
            sq = sq
        End If

' ************ 4
If Form6.Check4.Value Then
            sq = sq + Tbl + ".[ЖилаяПЛ], "
        Else
            sq = sq
        End If

' ************ 5
If Form6.Check5.Value Then
            sq = sq + Tbl + ".[Площадь], "
        Else
            sq = sq
       End If
       
  ' ************ 6
If Form6.Check6.Value Then
            sq = sq + Tbl + ".[Проживает], "
        Else
            sq = sq
       End If
       
' ************ 7
If Form6.Check7.Value Then
            sq = sq + Tbl + ".[Прописано], "
        Else
            sq = sq
       End If
       
 ' ************ 8
If Form6.Check8.Value Then
            sq = sq + Tbl + ".[Начисление], "
        Else
            sq = sq
       End If
       
 ' ************ 9
If Form6.Check9.Value Then
            sq = sq + Tbl + ".[Счет_затрат], "
        Else
            sq = sq
       End If
       
  ' ************ 18
If Form6.Check18.Value Then
            sq = sq + Tbl + ".[Сумма], "
        Else
            sq = sq
       End If
       
       ' ************ 10
If Form6.Check10.Value Then
            sq = sq + Tbl + ".[Кухня], "
        Else
            sq = sq
       End If
       
 ' ************ 11
If Form6.Check11.Value Then
            sq = sq + Tbl + ".[Ванная], "
        Else
            sq = sq
       End If
       
       
 ' ************ 12
If Form6.Check12.Value Then
            sq = sq + Tbl + ".[Коридор], "
        Else
            sq = sq
       End If
       
 ' ************ 13
If Form6.Check13.Value Then
            sq = sq + Tbl + ".[Туалет], "
        Else
            sq = sq
       End If
       
       
' ************ 14
If Form6.Check14.Value Then
            sq = sq + Tbl + ".[Балкон], "
        Else
            sq = sq
       End If
       
       
 ' ************ 15
If Form6.Check15.Value Then
            sq = sq + Tbl + ".[Этаж], "
        Else
            sq = sq
       End If
       
' ************ 16
If Form6.Check16.Value Then
            sq = sq + Tbl + ".[ДатаПрописки], "
        Else
            sq = sq
       End If
       
' ************ 17
If Form6.Check17.Value Then
            sq = sq + Tbl + ".[Прописка], "
        Else
            sq = sq
       End If
   
'*********************************************
        sq = Left(sq, Len(sq) - 2) + " "
sq = sq + "FROM " + Tbl

End Sub

