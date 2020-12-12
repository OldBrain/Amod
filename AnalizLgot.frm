VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AnalizLgot1 
   Caption         =   "Анализ "
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form7"
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Выход"
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
      Left            =   9360
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Убрать левую колонку "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7560
      Width           =   2055
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   13455
      _cx             =   23733
      _cy             =   9975
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"AnalizLgot.frx":0000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Command11 
         Caption         =   "XL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   15
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Настроить"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "+"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Расчитать"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         Picture         =   "AnalizLgot.frx":00DE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Объеденить"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Развернуть"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Отмена расчета"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   3720
         X2              =   3720
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   3720
         X2              =   3720
         Y1              =   0
         Y2              =   360
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   5730
      TabIndex        =   14
      Top             =   1080
      Width           =   75
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2160
      X2              =   2400
      Y1              =   960
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2160
      X2              =   2400
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Перед печатью  отчета "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"AnalizLgot.frx":0610
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   11055
   End
End
Attribute VB_Name = "AnalizLgot1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_kat As ADODB.Recordset
Dim CMB1 As ADODB.Recordset

Dim TheConn As ADODB.Connection
Dim Q, F As String
Public G As Integer, Titl As String


'Ur- уровень группировки
Private Sub Command1_Click()
On Error Resume Next
FG1.Subtotal flexSTSum, 1, 4, FG1.Cols, vbBlue, vbWhite, False, "И того"
FG1.Subtotal flexSTSum, 1, 5, FG1.Cols, vbBlue, vbWhite, True
FG1.Subtotal flexSTSum, 1, 6, FG1.Cols, vbBlue, vbWhite, True
FG1.Subtotal flexSTSum, 1, 7, FG1.Cols, vbBlue, vbWhite, True
FG1.Subtotal flexSTSum, 1, 8, FG1.Cols, vbBlue, vbWhite, True
End Sub

Private Sub Command10_Click()
FG1.Cols = FG1.Cols - 1
End Sub

Private Sub Command11_Click()
Pod.Show
Pod.Label1 = "Подождите идет экспорт данных в XL"

For i = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = i
 For j = 1 To 1000
    Next j
   Next

FG1.Subtotal flexSTClear
For i = 250 To 500
    Pod.ProgressBar1.Value = i
 For j = 1 To 1000
    Next j
   Next

FG1.DataRefresh
For i = 500 To 750
    Pod.ProgressBar1.Value = i
    
 For j = 1 To 1000
    Next j
   Next

ВыводВExel
For i = 750 To 1000
    Pod.ProgressBar1.Value = i
    
 For j = 1 To 1000
    Next j
   Next

Unload Pod

End Sub



Private Sub Command2_Click()
PrintW.Show
     With PrintW.VP
        PrintW.VP.StartDoc
        .FontSize = 12
        .Paragraph = Label3 + vbNewLine + "_________________________________________________________________"
        .Paragraph = ""
        
        .FontSize = 8
        .RenderControl = FG1.hwnd
        .EndDoc
        
       End With
       
       'With PrintW.VP
        '.StartDoc
         '   .SpaceAfter = 200
          '  .TextAlign = taJustTop
           ' .FontSize = 18
            '.Paragraph = "Hello, World."
            '.FontSize = 12
            '.Paragraph = "This is a VSPrinter document. You can set " & _
             '            "the margins by dragging them with the mouse."
            '.Paragraph = "Right now the margins are set as follows:"
            '.EndDoc
    'End With
    
    
End Sub

Private Sub Command3_Click()
FG1.MergeCells = flexMergeRestrictAll
 FG1.MergeCol(-1) = True
 FG1.MergeCol(FG1.Cols - 1) = False
 
End Sub

Private Sub Command4_Click()
FG1.MergeCells = flexMergeNever


End Sub

Private Sub Command5_Click()
FG1.Subtotal flexSTClear
FG1.DataRefresh
End Sub

Private Sub Command6_Click()
If FG1.Font.Size >= 1 Then FG1.Font.Size = FG1.Font.Size - 1
FG1.Refresh
End Sub

Private Sub Command7_Click()
Dim TMP As Double
TMP = FG1.Font.Size
FG1.Font.Size = FG1.Font.Size + 1
If FG1.Font.Size = TMP Then FG1.Font.Size = FG1.Font.Size + 2
FG1.Refresh

End Sub

Private Sub Command8_Click()
MainMenu.Enabled = True
MainMenu.Show
Unload Me
End Sub

Private Sub Command9_Click()
Dim i As Integer
Ur1 = 10
Uroven.Show
'AnalizLgot.Enabled = False



End Sub

Private Sub FG1_AfterMoveColumn(ByVal Col As Long, Position As Long)
' sort the data from first to last column
     FG1.Select 1, 0, 1, FG1.Cols - 1
     FG1.Sort = flexSortGenericAscending
     FG1.Select 1, 0
End Sub



Private Sub Form_Activate()
Pod.Hide
Label3 = Titl
End Sub

Private Sub Form_Load()



 Set TheConn = New ADODB.Connection
  TheConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  TheConn.Open "data/Kvartplata.mdb"
    
Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = TheConn
 
Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic



Rs_kat.Open (Reports.sq)
Pod.Show
Pod.Label1 = "Подождите, идет формирование отчета!"
Pod.Refresh


'For i = Pod.ProgressBar1.min To Pod.ProgressBar1.Max

For i = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = i
    
 For j = 1 To 1000
    Next j
   Next
'DoEvents
    




'Rs_kat.Filter = "[FAM] <> A"


Set FG1.DataSource = Rs_kat
 
For i = 250 To 300
    Pod.ProgressBar1.Value = i
    
  For j = 1 To 100000
    Next j
   Next


FG1.AllowUserResizing = flexResizeBoth

FG1.Cols = G
 '       FG1.FixedCols = 0
        'FG1.GridLinesFixed = flexGridExplorer
        'FG1.AllowUserResizing = flexResizeBoth
        FG1.ExplorerBar = flexExMove
        FG1.Editable = 2
        
  '      FG1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove

 FG1.MergeCells = flexMergeRestrictAll
 FG1.MergeCol(-1) = True
 FG1.MergeCol(FG1.Cols - 1) = False
        ' установите слияние ячейки (все колонны)
        'fg.MergeCells = flexMergeRestrictAll
        FG1.MergeCol(-1) = True



'Группировка
FG1.MergeCells = flexMergeRestrictAll
FG1.MergeCol(-1) = True
FG1.Refresh

FG1.Sort = flexSortGenericAscending
FG1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
If ur <> 0 Then NN = 500 Else NN = 1000
For i = 300 To NN
    Pod.ProgressBar1.Value = i
     For j = 1 To 100000
    Next j
   Next



End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Uroven
Unload Pod
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "Save"
            'ToDo: Add 'Save' button code.
           Unload Me
           Reports.Show
               End Select
End Sub

Public Sub Об(ur As Integer)


On Error GoTo er1
'MsgBox (Str(ur))


         '*************расчет I-ой колонки ***************************
         
     If ur = 0 Then Exit Sub
     If ur = 1 Then
     
     ' suspend repainting to get more speed
        FG1.Redraw = False
        FG1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        FG1.Select 1, 0, 1, FG1.Cols - 1
        FG1.Sort = flexSortGenericAscending
        FG1.Select 1, 0
        ' calculate subtotals
        FG1.Subtotal flexSTClear

     
        For i = 2 To FG1.Cols - 1
        FG1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        FG1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
         Next
          ' autosize
        FG1.AutoSize 0, FG1.Cols - 1, , 300
        ' turn repainting back on
        FG1.OutlineBar = flexOutlineBarComplete
        FG1.Redraw = True
        Unload Uroven
     End If
        
     If ur = 2 Then
     ' suspend repainting to get more speed
        FG1.Redraw = False
        FG1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        FG1.Select 1, 0, 1, FG1.Cols - 1
        FG1.Sort = flexSortGenericAscending
        FG1.Select 1, 0
        ' calculate subtotals
        FG1.Subtotal flexSTClear
     
     
     
     For i = 2 To FG1.Cols - 1
        'fg1.Subtotal flexSTSum, -1, i, , RGB(200, 255, 200), vbBlack, True, "ИТОГ:"
        FG1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        FG1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
        FG1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, True
        
    Next
     ' autosize
        FG1.AutoSize 0, FG1.Cols - 1, , 300
        ' turn repainting back on
        FG1.OutlineBar = flexOutlineBarComplete
        FG1.Redraw = True
        Unload Uroven
      End If
        
        
     If ur = 3 Then
     ' suspend repainting to get more speed
        FG1.Redraw = False
        FG1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        FG1.Select 1, 0, 1, FG1.Cols - 1
        FG1.Sort = flexSortGenericAscending
        FG1.Select 1, 0
       
        
        ' calculate subtotals
        FG1.Subtotal flexSTClear
     
     For i = 2 To FG1.Cols - 1
        FG1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        FG1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, False, "И того:"
        FG1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, False
        FG1.Subtotal flexSTSum, 2, i, , RGB(200, 250, 200), vbBlack, False
       Next
        ' autosize
        FG1.AutoSize 0, FG1.Cols - 1, , 300
        ' turn repainting back on
        FG1.OutlineBar = flexOutlineBarComplete
        FG1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If

 If ur = 4 Then
     ' suspend repainting to get more speed
        FG1.Redraw = False
        FG1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        FG1.Select 1, 0, 1, FG1.Cols - 1
        FG1.Sort = flexSortGenericAscending
        FG1.Select 1, 0
        ' calculate subtotals
        FG1.Subtotal flexSTClear
     
     
     
     
     For i = 2 To FG1.Cols - 1
        FG1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        FG1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
        FG1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, False
        FG1.Subtotal flexSTSum, 2, i, , RGB(200, 250, 200), vbBlack, False
        FG1.Subtotal flexSTSum, 3, i, , RGB(100, 200, 200), vbBlack, False
           Next
        ' autosize
        FG1.AutoSize 0, FG1.Cols - 1, , 300
        ' turn repainting back on
        FG1.OutlineBar = flexOutlineBarComplete
        FG1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If

If ur = 5 Then
     ' suspend repainting to get more speed
        FG1.Redraw = False
        FG1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        FG1.Select 1, 0, 1, FG1.Cols - 1
        FG1.Sort = flexSortGenericAscending
        FG1.Select 1, 0
        ' calculate subtotals
        FG1.Subtotal flexSTClear
     
     
     
     
     For i = 2 To FG1.Cols - 1
        FG1.Subtotal flexSTSum, -1, i, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        FG1.Subtotal flexSTSum, 0, i, , RGB(300, 300, 250), vbBlack, True
        FG1.Subtotal flexSTSum, 1, i, , RGB(220, 380, 250), vbBlack, False
        FG1.Subtotal flexSTSum, 2, i, , RGB(200, 250, 200), vbBlack, False
        FG1.Subtotal flexSTSum, 3, i, , RGB(100, 200, 200), vbBlack, False
        FG1.Subtotal flexSTSum, 4, i, , RGB(300, 100, 200), vbBlack, False
       Next
        ' autosize
        FG1.AutoSize 0, FG1.Cols - 1, , 300
        ' turn repainting back on
        FG1.OutlineBar = flexOutlineBarComplete
        FG1.Redraw = True
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
MsgBox "Нет данных для отчета с выбранными параметрами " + Analizlgot.Caption
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

Sub ВыводВExel()
   Const НачСтрока = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim i As Long, j As Long, k As Long, rДанные As String
   Dim v As Variant
   
   
'rs.CursorType = adOpenStatic
'rs.LockType = adLockReadOnly



   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
'   Set rs = Rs_kat.Clone

 'Set rs = Rs_kat
  ' rs.Filter = Rs_kat.Filter
   'rs.Sort = Rs_kat.Sort
 ' k = FG1.Rows - 1
  ' Rs_kat.MoveLast
'   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(k) & Rs_kat.RecordCount + НачСтрока
   
   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(FG1.Cols - 1) & FG1.Rows + НачСтрока
   ReDim v(FG1.Rows, FG1.Cols) 'Забыл указать
'   If rs.RecordCount > 0 Then

   'If Rs_kat.RecordCount > 0 Then
   If FG1.Rows > 0 Then
    '  Rs_kat.MoveFirst
      'i = 0
      'Do Until Rs_kat.EOF
         For co = 1 To FG1.Cols - 1
         For Rw = 0 To FG1.Rows - 1
             'v(i, j) = Rs_kat.Fields(j).Value
             v(Rw, co) = FG1.TextMatrix(Rw, co)
             
         Next Rw
         Next co
         'Rs_kat.MoveNext
      'Loop
      ex1.Visible = True   'Еще забыл
      
      ws.Range(rДанные) = v
      
 End If
End Sub

Function XCol_(ByVal Column_ As Long) As String
    If (Column_ < 0) Then Column_ = 0
    If (Column_ < 26) Then
        XCol_ = Chr(Column_ + Asc("A"))
    ElseIf (Column_ < 676) Then
        XCol_ = Chr((Column_ \ 26) + Asc("A") - 1) & Chr((Column_ Mod 26) + Asc("A"))
    Else
        XCol_ = "ZZ"
    End If
End Function


