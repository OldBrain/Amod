VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DomTip 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Тарифы"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7845
   ClipControls    =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Удалить <F8>"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Добавить<F4>"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ok"
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
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   9135
      _cx             =   16113
      _cy             =   9551
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"DomTip.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   2
      ExplorerBar     =   7
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
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Т А Р И Ф Ы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9135
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Добавить 
         Caption         =   "Добавить"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Удалить"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Закрыть 
         Caption         =   "Закрыть"
      End
   End
End
Attribute VB_Name = "DomTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cl As String
Dim Rs_kat As ADODB.Recordset
Dim TheConn As ADODB.Connection
'Dim Combo_Conn As ADODB.Connection
Dim Combo_RS As ADODB.Recordset



Private Sub DataList1_Click()
DataList1.Refresh
End Sub

Private Sub Calendar1_DblClick()
Calendar1.GridCellEffect

End Sub

Private Sub Command1_Click()
Rs_kat.UpdateBatch
MainMenu.Show
Tarif.Hide
End Sub

Private Sub Command2_Click()

Dim N, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
N = 0
Rs_kat.MoveFirst
Do While Not Rs_kat.EOF
If Rs_kat("Код").Value = "" Then
Rs_kat.Delete
Rs_kat.MoveFirst
End If
N1 = Rs_kat("Код").Value
If N1 > N Then N = N1
Rs_kat.MoveNext
Loop

Rs_kat.AddNew
Rs_kat("Код") = N + 1
Rs_kat("KATEGOR") = "Новая запись"
Rs_kat.UpdateBatch
FG1.DataRefresh
Rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With Rs_kat
DelItem = FG1.TextMatrix(FG1.Row, 1)
If MsgBox("Вы хотите удалить " + FG1.TextMatrix(FG1.Row, 2) + " " + FG1.TextMatrix(FG1.Row, 3) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If Rs_kat("Код") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
FG1.DataRefresh
If .EOF Then .MoveLast
End If
End With

End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
' Tarif.FG1.Cell(flexcpText) = MainForm.TMP1
Rs_kat.UpdateBatch
End Sub

Private Sub FG1_Click()


If (FG1.TextMatrix(0, FG1.Col)) = "Тип" Then
cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
cl = cl + Combo_RS("Name_Kategor") + "|"
Combo_RS.MoveNext
Loop
FG1.ComboList = cl
Else: FG1.ComboList = ""
End If
End Sub

Private Sub FG1_GotFocus()
FG1.ComboList = ""
End Sub


Private Sub FG1_RowColChange()
FG1.Refresh
End Sub

Private Sub FG1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


If (FG1.TextMatrix(0, FG1.Col)) = "Тип" Then
cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
cl = cl + Combo_RS("Name_Kategor") + "|"
Combo_RS.MoveNext
Loop
FG1.ComboList = cl
Else: FG1.ComboList = ""
End If
End Sub

Private Sub Form_Load()
FG1.ComboList = ""
 FG1.Editable = False

' open connection
   Set TheConn = New ADODB.Connection
  TheConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  TheConn.Open "data/Kvartplata.mdb"
  
 
  
Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = TheConn
Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic
Rs_kat.Open "Socmin"
Set FG1.DataSource = Rs_kat

Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = TheConn
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
Combo_RS.Open "Kategor"

'FG1.ColComboList(2) = "..."  ' date picker popup


' правопреемник recordset в сетку
   
   FG1.FocusRect = flexFocusSolid
    FG1.Editable = True
    FG1.DataMode = flexDMBound
    
    FG1.AutoSearch = flexSearchFromCursor
    FG1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs_kat.Close
TheConn.Close
End Sub

Private Sub Добавить_Click()
Command2_Click
End Sub

Private Sub Закрыть_Click()
Command1_Click
End Sub

Private Sub Удалить_Click()
Command3_Click
End Sub
