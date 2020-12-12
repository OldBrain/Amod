VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Jilec 
   Caption         =   "Справочник жильцов"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9960
   LinkTopic       =   "Form7"
   ScaleHeight     =   6945
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9495
      _cx             =   16748
      _cy             =   10610
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
End
Attribute VB_Name = "Jilec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection

Private Sub DataList1_Click()
DataList1.Refresh
End Sub

Private Sub Command1_Click()
Rs_kat.UpdateBatch
Doma.Hide
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
Rs_kat("NAIM_KLS") = "Новый адрес"
Rs_kat.UpdateBatch
FG1.DataRefresh
Rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With Rs_kat
DelItem = FG1.TextMatrix(FG1.Row, 1)
If MsgBox("Вы хотите удалить " + FG1.TextMatrix(FG1.Row, 2) + "?", vbYesNo) = vbYes Then
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
Rs_kat.UpdateBatch

End Sub

Private Sub FG1_Click()
'MsgBox (FG1.Cell(flexcpText))
'MsgBox (FG1.TextMatrix(FG1.Row, 1))
End Sub

Private Sub FG1_RowColChange()
FG1.Refresh
End Sub

Private Sub Form_Load()

 FG1.Editable = False

' open connection
 '  Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"
    
Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = mconn
 
Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic
Rs_kat.Open "MainOccupant"
Set FG1.DataSource = Rs_kat


' правопреемник recordset в сетку
   
   FG1.FocusRect = 3
    'flexFocusSolid
    FG1.Editable = True
    FG1.DataMode = flexDMBound
    
    FG1.AutoSearch = flexSearchFromCursor
    FG1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs_kat.Close
mconn.Close
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

