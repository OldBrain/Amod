VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Analizlgot 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
   FillStyle       =   0  'Solid
   HasDC           =   0   'False
   Icon            =   "xpForm.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   680
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   960
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin KvPay.xpcmdbutton Dol 
      Height          =   252
      Left            =   1560
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   445
      Caption         =   "Печать уведомлений для должников"
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
   Begin VB.CommandButton Command11 
      BackColor       =   &H00BDC6BB&
      Caption         =   "XL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00BDC6BB&
      Height          =   315
      Left            =   8760
      Picture         =   "xpForm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Настройка"
      Height          =   315
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Отм.расч"
      Height          =   315
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Расчитать"
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Развернуть"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Объеденить"
      Height          =   315
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Image1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Выход"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Пересчитать сальдо конечное"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Проставить сальдо начальное"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   7020
      _cx             =   12382
      _cy             =   9975
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      FormatString    =   $"xpForm.frx":0126
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   195
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5640
      Picture         =   "xpForm.frx":0204
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   6240
      Picture         =   "xpForm.frx":094E
      Top             =   120
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   5880
      Picture         =   "xpForm.frx":1098
      Top             =   120
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
      Left            =   6840
      Picture         =   "xpForm.frx":17E2
      Top             =   120
      Width           =   156
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1890
   End
End
Attribute VB_Name = "Analizlgot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************
' * WARNING                                               *
' * =======                                               *
' *                                                       *
' * This code relies heavily on the Z-Index (arrangement, *
' * send to back etc.) of the elements that make up the   *
' * skinned window.  Therefore, if you use this in your   *
' * own programs, you might have to fiddle about sending  *
' * things to back, front and the like before it works.   *
' *********************************************************
Public rs_kat As ADODB.Recordset
Public Vid As String
Public Ok As Double

' Если отчет собирает должников то истина
'Public Dolg As Boolean


'Dim Saldo As ADODB.Recordset
Dim CMB1 As ADODB.Recordset
Dim PRS As ADODB.Recordset
'Dim Mconn As ADODB.Connection
Dim Q, F As String
Dim fil As Integer
Public StrSQL As String
Public G As Integer, Titl As String
'*****************************************
Dim Temp
Dim flgResize As Boolean
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI

Private Sub Combo1_GotFocus()
Command6.Visible = True
SendKeys "{F4}"
Combo1.SetFocus
End Sub

Private Sub Combo1_LostFocus()
Command6.Visible = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
fg1.Subtotal flexSTSum, 1, 4, fg1.Cols, vbBlue, vbWhite, False, "И того"
fg1.Subtotal flexSTSum, 1, 5, fg1.Cols, vbBlue, vbWhite, True
fg1.Subtotal flexSTSum, 1, 6, fg1.Cols, vbBlue, vbWhite, True
fg1.Subtotal flexSTSum, 1, 7, fg1.Cols, vbBlue, vbWhite, True
fg1.Subtotal flexSTSum, 1, 8, fg1.Cols, vbBlue, vbWhite, True
End Sub

Private Sub Command11_1_Click()

End Sub

Private Sub Command11_Click()
'MsgBox ("Операция временно заблокирована")
'Exit Sub

Pod.Enabled = True
Pod.Label1 = "Подождите идет экспорт данных в XL"

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

fg1.Subtotal flexSTClear
For I = 250 To 500
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

fg1.DataRefresh
For I = 500 To 750
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

ВыводВExel
For I = 750 To 1000
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
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
        .RenderControl = fg1.hwnd
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


Private Sub Command4_Click()
fg1.MergeCells = flexMergeNever
End Sub

Private Sub Command3_Click()
fg1.MergeCells = flexMergeRestrictAll
 fg1.MergeCol(-1) = True
 fg1.MergeCol(fg1.Cols - 1) = False
 
 End Sub

Private Sub Command5_Click()
fg1.Subtotal flexSTClear
fg1.DataRefresh
End Sub



Private Sub Command6_Click()
Jdite.Show
Jdite.Label1 = "Добавляю недостающее сальдо"

On Error GoTo Er
Mconn.Execute ("INSERT INTO Adding ( KodKv, KodKat, SaldoN, KodN, NameN, Formula, Kol, FormulaB, Tarif, Socmin, Propis ) SELECT SAбезAdd.KodKV, SAбезAdd.KodKat, SAбезAdd.SK, 999 AS Выражение1," + Chr(34) + "Сальдо" + Chr(34) + " AS Выражение2, " + Chr(34) + "SummaI" + Chr(34) + " AS Выражение3, 1 AS Выражение4, " + Chr(34) + "SummaI" + Chr(34) + " AS Выражение5, 0 AS Выражение6, 0 AS Выражение7, 0 AS Выражение8 FROM SAбезAdd")
Er:
Jdite.Label1 = "Проставляю  правильное сальдо"
Mconn.Execute ("UPDATE Рас_с_нач INNER JOIN ADDING ON (Рас_с_нач.KodKV = ADDING.KodKv) AND (Рас_с_нач.KodKat = ADDING.KodKat) SET ADDING.SaldoN = Рас_с_нач![Сальдо  на конец]")
Unload Jdite
MsgBox "Сальдо проставлено успешно"
Command7_Click
'Unload Me
End Sub

Private Sub Command7_Click()
'Dim rsRash As ADODB.Recordset
'Set rsRash = New ADODB.Recordset

K = 0
Me.Enabled = False


Pod.Show
Pod.Enabled = False

Pod.Label1.Caption = "П О Д О Ж Д И Т Е"

Mconn.Execute ("DELETE PRS.* FROM PRS")
Mconn.Execute ("INSERT INTO PRS ( OLDNUM, FAM, IM, OT, KodKv, KodKat, СН, НЧ, СК, Отклонение ) SELECT MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, [Проверка Сальдо1].KodKv, [Проверка Сальдо1].KodKat, [Проверка Сальдо1].СН, [Проверка Сальдо1].НЧ, [Проверка Сальдо1].СК, [Проверка Сальдо1]!СК-([Проверка Сальдо1]!СН+[Проверка Сальдо1]!НЧ) AS Отклонение FROM [Проверка Сальдо1] INNER JOIN MainOccupant ON [Проверка Сальдо1].KodKv = MainOccupant.Numer WHERE (((Round([Проверка Сальдо1]![СК]-([Проверка Сальдо1]![СН]+[Проверка Сальдо1]![НЧ]),2))<>0))")
Set PRS = New ADODB.Recordset
'PRS.Open ("SELECT PRS.KodKv, PRS.FAM FROM PRS"), Mconn, adOpenKeyset
PRS.Open ("PRS"), Mconn
Set fg1.DataSource = PRS

Pod.Label1.Caption = "П О Д О Ж Д И Т Е"
On Error GoTo nol
PRS.MoveFirst
Do While Not PRS.EOF
K = K + 1
PRS.MoveNext
Loop



PRS.MoveFirst
Pod.ProgressBar1.Value = 1
Pod.ProgressBar1.Max = K + 1
Do While Not PRS.EOF

Pod.Label1.Caption = "П О Д О Ж Д И Т Е," + vbNewLine + "пересчитываю лиц.сч. №" + Str(PRS("Kodkv")) + " " + PRS("FAM")
Pod.Label1.Refresh
'Расчет сальдо и количества
MainForm.КоличествоСальдо Str(PRS("Kodkv"))
MainForm.RSaldoK Str(PRS("Kodkv"))


Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1

PRS.MoveNext
Loop
nol:
Pod.Label1 = "Расчет сальдо на конец периода выполнен"
Unload Pod
MainMenu.Enabled = False
PRS.Close
Set PRS = Nothing
Me.Enabled = True
Me.Show


End Sub

Private Sub Command8_Click()
Отчет "lc.doc"
End Sub

Private Sub Command9_Click()
Dim I As Integer
Ur1 = 10
Uroven.Show
End Sub


Private Sub Dol_Click()


For I = 3 To fg1.Rows - 1
' MsgBox (FG1.Rows)
'MsgBox (FG1.TextMatrix(I, 1))

Sud Val(Str(fg1.TextMatrix(I, 1)))

Next I
End Sub

Private Sub FG1_Click()
'MsgBox FG1.TextMatrix(4, 1)
On Error GoTo osh
If Vid = "ОбВд" Then
Label1.Visible = True
'Label2.Visible = True

Ok = Val(Str(fg1.TextMatrix(1, 3))) + Val(Str(fg1.TextMatrix(1, 4))) - Val(Str(fg1.TextMatrix(1, 5))) - Val(Str(fg1.TextMatrix(1, 6))) - Val(Str(fg1.TextMatrix(1, 7)))

Label1.Caption = fg1.TextMatrix(1, 3) + " + " + fg1.TextMatrix(1, 4) + " - " + fg1.TextMatrix(1, 5) + " - " + fg1.TextMatrix(1, 6) + " - " + fg1.TextMatrix(1, 7) + " = " + Str(Round(Ok, 2))
'Label2.Caption = Str(Ok)

If Round(Ok, 2) <> 0 Then
Command6.Visible = True
Command7.Visible = True

'FG1.Cell(flexcpFontBold, 2, 1, 2, FG1.Cols - 1) = True
'FG1.Cell(flexcpBackColor, 2, 1, 2, FG1.Cols - 1) = vbRed

Else

osh:
If Err.Number <> 0 Then
MsgBox Err.Description
Err.Clear
End If
End If



End If
End Sub

Private Sub Form_Activate()
 Unload Pod
Label3 = Titl
If Analizlgot.Titl = "Расхождение сальдо на начало периода" Then
Command6.Visible = True
Command7.Visible = True
End If
End Sub

Private Sub Image1_Click()
Set rs_kat = Nothing
Set fg1.DataSource = Nothing


MainMenu.Enabled = True
Unload Pod
MainMenu.Enabled = True

' MainMenu.Show

Unload Reports
Unload Me
Unload Analizlgot
End Sub

Private Sub Form_Load()
    MakeWindow Me, True
    Set fg1.DataSource = Nothing
    
Label3 = Titl
'Dol.Visible = Dolg
    
    
On Error GoTo erRep
    
    'AlwaysOnTop Me, True
' Make the Maximize/Restore button have the Maximize image
'   imgTitleMaxRestore.Picture = imgTitleMaximize.Picture
   fg1.Width = Me.Width / 15.40107
   fg1.Height = Me.Height / 20
   Image1.Top = Me.Height / 16.16477
   Image1.Left = 3
   
   
   Command1.Top = Image1.Top
   Command3.Top = Image1.Top
   Command4.Top = Image1.Top
   Command5.Top = Image1.Top
   Command9.Top = Image1.Top
   Command2.Top = Image1.Top
   Command11.Top = Image1.Top
 '***********************
' Коннект
'  Set Mconn = New ADODB.Connection
 ' Mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' Mconn.Open "data/Kvartplata.mdb"
    
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
 
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic

If Reports.sq <> "" Then StrSQL = Reports.sq

If StrSQL <> "" Then rs_kat.Open (StrSQL)
Pod.Show
Pod.Label1 = "Подождите, идет формирование отчета!"
Pod.Refresh

'For i = Pod.ProgressBar1.min To Pod.ProgressBar1.Max

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next
'DoEvents
    




'Rs_kat.Filter = "[FAM] <> A"


Set fg1.DataSource = rs_kat
 
For I = 250 To 300
    Pod.ProgressBar1.Value = I
    
  For J = 1 To 100000
    Next J
   Next


fg1.AllowUserResizing = flexResizeBoth

fg1.Cols = G
 '       FG1.FixedCols = 0
        'FG1.GridLinesFixed = flexGridExplorer
        'FG1.AllowUserResizing = flexResizeBoth
        fg1.ExplorerBar = flexExMove
        
        'FG1.Editable = 2
        
  '      FG1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove

 fg1.MergeCells = flexMergeRestrictAll
 fg1.MergeCol(-1) = True
 fg1.MergeCol(fg1.Cols - 1) = False
        ' установите слияние ячейки (все колонны)
        'fg.MergeCells = flexMergeRestrictAll
        fg1.MergeCol(-1) = True






'Группировка
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.Refresh

fg1.Sort = flexSortGenericAscending

fg1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
If ur <> 0 Then NN = 500 Else NN = 1000

For I = 300 To NN
   Pod.ProgressBar1.Value = I
    For J = 1 To 100000 Step 100
  Next J
Next





'FG1.RowHeight(0) = 700
fg1.RowHeight(0) = 500

fg1.WordWrap = True
fg1.Cell(flexcpAlignment, 0, 0, 0, fg1.Cols - 1) = flexAlignCenterCenter


erRep:
If Err.Number <> 0 Then MsgBox Err.Description


'Pod.ProgressBar1.Max = FG1.Rows + 1
'For rw = 1 To FG1.Rows - 1
'Pod.ProgressBar1.Value = rw
'FG1.TextMatrix(rw, 0) = rw
'Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
Vid = ""
Unload Uroven
Unload Pod
End Sub

 '*********************
   


Private Sub Form_Resize()
fg1.Width = Me.Width / 15.40107
   fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
Command4.Top = Image1.Top
Command1.Top = Image1.Top
Command5.Top = Image1.Top
Command9.Top = Image1.Top
Command2.Top = Image1.Top
Command11.Top = Image1.Top

Command3.Left = Image1.Left + Image1.Width
Command4.Left = Image1.Left + Image1.Width + Command3.Width
Command5.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width
Command1.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width
Command9.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width + Command1.Width
Command2.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width + Command1.Width + Command9.Width
Command11.Left = Image1.Left + Image1.Width + Command3.Width + Command4.Width + Command5.Width + Command1.Width + Command9.Width + Command2.Width
End Sub


Private Sub Label2_Click()

End Sub

Private Sub imgTitleHelp_Click()
About.Show

End Sub

Private Sub lblTitle_DblClick()
    ChangeState Me
End Sub

Private Sub imgTitleMain_DblClick()
    ChangeState Me
    
End Sub

Public Sub Об(ur As Integer)


On Error GoTo er1
'MsgBox (Str(ur))


         '*************расчет I-ой колонки ***************************
         
     If ur = 0 Then Exit Sub
     If ur = 1 Then
     
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear

     
        For I = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, I, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, I, , RGB(300, 300, 250), vbBlack, True
         Next
          ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
     End If
        
     If ur = 2 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     
     
     For I = 2 To fg1.Cols - 1
        'fg1.Subtotal flexSTSum, -1, i, , RGB(200, 255, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, -1, I, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, I, , RGB(300, 300, 250), vbBlack, True
        fg1.Subtotal flexSTSum, 1, I, , RGB(220, 380, 250), vbBlack, True
        
    Next
     ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
        
        
     If ur = 3 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
       
        
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     For I = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, I, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, I, , RGB(300, 300, 250), vbBlack, False, "И того:"
        fg1.Subtotal flexSTSum, 1, I, , RGB(220, 380, 250), vbBlack, False
        fg1.Subtotal flexSTSum, 2, I, , RGB(200, 250, 200), vbBlack, False
       Next
        ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If

 If ur = 4 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     
     
     
     For I = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, I, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, I, , RGB(300, 300, 250), vbBlack, True
        fg1.Subtotal flexSTSum, 1, I, , RGB(220, 380, 250), vbBlack, False
        fg1.Subtotal flexSTSum, 2, I, , RGB(200, 250, 200), vbBlack, False
        fg1.Subtotal flexSTSum, 3, I, , RGB(100, 200, 200), vbBlack, False
           Next
        ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If

If ur = 5 Then
     ' suspend repainting to get more speed
        fg1.Redraw = False
        fg1.MergeCells = flexMergeRestrictAll
        ' sort the data from first to last column
        fg1.Select 1, 0, 1, fg1.Cols - 1
        fg1.Sort = flexSortGenericAscending
        fg1.Select 1, 0
        ' calculate subtotals
        fg1.Subtotal flexSTClear
     
     
     
     
     For I = 2 To fg1.Cols - 1
        fg1.Subtotal flexSTSum, -1, I, , RGB(250, 250, 200), vbBlack, True, "ИТОГ:"
        fg1.Subtotal flexSTSum, 0, I, , RGB(300, 300, 250), vbBlack, True
        fg1.Subtotal flexSTSum, 1, I, , RGB(220, 380, 250), vbBlack, False
        fg1.Subtotal flexSTSum, 2, I, , RGB(200, 250, 200), vbBlack, False
        fg1.Subtotal flexSTSum, 3, I, , RGB(100, 200, 200), vbBlack, False
        fg1.Subtotal flexSTSum, 4, I, , RGB(300, 100, 200), vbBlack, False
       Next
        ' autosize
        fg1.AutoSize 0, fg1.Cols - 1, , 300
        ' turn repainting back on
        fg1.OutlineBar = flexOutlineBarComplete
        fg1.Redraw = True
        Unload Uroven
      End If
      
     If ur = 10 Then
     Exit Sub
     Unload Uroven
     End If
     
     For I = 500 To 1000
    Pod.ProgressBar1.Value = I
    For J = 1 To 100000
    Next J
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
   Dim I As Long, J As Long, K As Long, rДанные As String
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
   
   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(fg1.Cols - 1) & fg1.Rows + НачСтрока
   ReDim v(fg1.Rows, fg1.Cols) 'Забыл указать
'   If rs.RecordCount > 0 Then

   'If Rs_kat.RecordCount > 0 Then
   If fg1.Rows > 0 Then
    '  Rs_kat.MoveFirst
      'i = 0
      'Do Until Rs_kat.EOF
         For co = 1 To fg1.Cols - 1
         For rw = 0 To fg1.Rows - 1
             'v(i, j) = Rs_kat.Fields(j).Value
             v(rw, co) = fg1.TextMatrix(rw, co)
             
         Next rw
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



Private Sub Sud(KeyNum As String)

Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'объявляем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim rsSud As ADODB.Recordset



Set rsSud = New ADODB.Recordset
Set rsSud.ActiveConnection = Mconn
 
rsSud.CursorType = adOpenForwardOnly
rsSud.LockType = adLockBatchOptimistic

rsSud.Open ("SELECT MainOccupant.Numer, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.NLODGERF, MainOccupant.COMSPACE FROM KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((MainOccupant.Numer)=" + KeyNum + "))")

'' MsgBox (rsSud("FAM") + " " + rsSud("IM") + " " + rsSud("OT"))

Dolg = Round(fg1.TextMatrix(3, 6), 2)
'FormDolg.Text1 = Dolg

'FormDolg.Show 1



If Dolg = -369.8985231 Then Exit Sub


nameRP = "Sud"

'//////////////////////////////////////////////
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = True

'создаём новый документ в Word-e
'Set DocWord = WordApp.Documents.Add

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его



DocWord.Activate

'сохраняем временный документ
On Error GoTo est
If Err.Number <> 5356 Then
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)

est:
 End If
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))
DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
WordApp.Options.CheckSpellingAsYouType = False
Set DocWord = WordApp.Documents.Open(App.Path + "\Temp\" + nameRP + ".doc")



DocWord.Activate
Set TableWord = DocWord.Tables(1)
TableWord.Cell(5, 2).Range.Text = rsSud("FAM") + " " + rsSud("IM") + " " + rsSud("OT")


TableWord.Cell(2, 3).Range.Text = "АДРЕС"

'Долг
TableWord.Cell(8, 2).Range.Text = "" + MainForm.Label8 + " г."
TableWord.Cell(8, 3).Range.Text = "составляет"
TableWord.Cell(8, 4).Range.Text = Str(Dolg) + " руб."


'TableWord.Cell(8, 3).Range.Text = Label10


'Площадь, прописано и т.д.
'TableWord.Cell(11, 1).Range.Text = "Общ.пл.-" + FG1.TextMatrix(FG1.Row, 15) + "м*2 Прописано-" + FG1.TextMatrix(FG1.Row, 12) + "ч."


'DocWord.Tables(1).Rows.Add

 
'TableWord.Cell(15, 1).Range.Text = NumStr(Dolg, True)

'//////////////////////////////////////////

'Копируем таблицу
 '   Dim Tbl As Table
   ' Dim rng As Range
    
    
    With WordApp.ActiveDocument
 Set rng = .Paragraphs(.Paragraphs.Count).Range
 
 
 
'    Set rng = WordApp.ActiveDocument.Paragraphs(WordApp.ActiveDocument.Paragraphs.Count).Range
        
        
'Добавляем строку
'DocWord.Tables(1).Columns.Add 13
'DocWord.Tables(1).Rows.Add


K = 15

'Сальдо
DocWord.Tables(1).Rows.Add
If Val(Label10) >= 0 Then
TableWord.Cell(8, 2).Range.Text = "Долг на начало " + MainForm.Label8 + " г."
TableWord.Cell(8, 3).Range.Text = Label10

End If

If Val(Label10) < 0 Then
TableWord.Cell(K + I, 1).Range.Text = "Переплата на начало " + MainForm.Label8 + " г."
TableWord.Cell(K + I, 2).Range.Text = Label10
End If

'K = 16
'For I = 1 To FG1.Rows - 1

'DocWord.Tables(1).Rows.Add
'наим.платежа

'If FG1.TextMatrix(I, 23) <> "+" Then TableWord.Cell(K + I, 1).Range.Text = FG1.TextMatrix(I, 3)
'If FG1.TextMatrix(I, 23) = "+" Then TableWord.Cell(K + I, 1).Range.Text = FG1.TextMatrix(I, 3) + " (по тар = " + FG1.TextMatrix(I, 10) + "руб.)"
'Сумма
'TableWord.Cell(K + I, 2).Range.Text = FG1.TextMatrix(I, 18)
'Статус
'TableWord.Cell(K + I, 3).Range.Text = FG1.TextMatrix(I, 23)
'End If
'Next
        
        
        
        
        
        
        Set Tbl = .Tables(1)
    End With
    
    
'K = 14
If Val(Label10) <> 0 Then
DocWord.Tables(1).Rows.Add
'наим.платежа




'Сумма
'TableWord.Cell(14, 1).Range.Text = FG1.TextMatrix(FG1.Row, 1)
TableWord.Cell(K + 1, 1).Range.Text = "И ТОГО К ОПЛАТЕ:"
TableWord.Cell(K + 1, 2).Range.Text = Dolg
'K = 15
End If
    
    
    
    
       rng.ParagraphFormat.Alignment = wdAlignParagraphRight
       
       
       '**********************************
       'rng.InsertAfter NumStr(Dolg, True)
        
       
    
 '   Tbl.Range.Copy
    
    
    With rng
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
    
        .Collapse Direction:=wdCollapseEnd
        .Paste

 End With

End Sub

