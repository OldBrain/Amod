VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Socmin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Справочник категорий расчетов"
   ClientHeight    =   7824
   ClientLeft      =   156
   ClientTop       =   336
   ClientWidth     =   6756
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7824
   ScaleWidth      =   6756
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6756
      _ExtentX        =   11917
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OOFL1"
            ImageKey        =   "OOFL"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000000&
      Caption         =   "Удалить <F8>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000000&
      Caption         =   "Добавить<F4>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
      _cx             =   11668
      _cy             =   11668
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
      BackColor       =   -2147483634
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483634
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Socmin.frx":0000
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
      ExplorerBar     =   5
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2670
      Top             =   3675
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Socmin.frx":00D8
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Socmin.frx":01EA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Socmin.frx":02FC
            Key             =   "OOFL"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Справочник соцминимумов по категориям расчета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6495
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
Attribute VB_Name = "Socmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cl As String
Dim rs_kat As ADODB.Recordset
Dim TMP As ADODB.Recordset


'Dim mconn As ADODB.Connection
'Dim Combo_Conn As ADODB.Connection
Dim Combo_RS As ADODB.Recordset
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    'On Error Resume Next
    Select Case Button.KEY
        Case "New"
            'ToDo: Add 'New' button code.
            Command2_Click
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            Command3_Click
        Case "OOFL1"
            'ToDo: Add 'OOFL1' button code.
            Command1_Click
    End Select
End Sub




Private Sub DataList1_Click()
DataList1.Refresh
End Sub

Private Sub Calendar1_DblClick()
Calendar1.GridCellEffect

End Sub

Private Sub Command1_Click()

DoEvents

Jdite.Show
Jdite.Label1.Refresh

rs_kat.UpdateBatch

' Теперь заполняем нулями пустые соцминимумы для Adding
Mconn.Execute ("UPDATE Adding SET Adding.Socmin = 0 WHERE (((Adding.Socmin) Is Null))")
' Обнавляем соцминимумы для Adding
Mconn.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.KodKat = Socmin.KodKategor) AND (Adding.Propis = Socmin.koli) SET Adding.Socmin = [Socmin]![Value]+[Adding]![dop]")
' Теперь заполняем нулями пустые соцминимумы для Adding
Mconn.Execute ("UPDATE Adding SET Adding.Socmin = 0 WHERE (((Adding.Socmin) Is Null))")
Mconn.Execute ("UPDATE Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd SET tmp_lgota.Cocmin = [Adding]![Socmin]")

'Set tmp = New ADODB.Recordset
'Set tmp.ActiveConnection = mconn
'tmp.CursorType = adOpenForwardOnly
'tmp.LockType = adLockBatchOptimistic
'tmp.Open "OBN_COCMIN"
'tmp.Close


Unload Jdite
Unload Me

Sprav.Show
'Socmin.Hide

End Sub

Private Sub Command2_Click()

Dim n, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
n = 0
rs_kat.MoveFirst
Do While Not rs_kat.EOF
If rs_kat("Код").Value = "" Then
rs_kat.Delete
rs_kat.MoveFirst
End If
N1 = rs_kat("Код").Value
If N1 > n Then n = N1
rs_kat.MoveNext
Loop

rs_kat.AddNew
rs_kat("Код") = n + 1
rs_kat("Period") = MainForm.PeriodR
rs_kat("KATEGOR") = "Новая запись"
rs_kat.UpdateBatch
Fg1.DataRefresh
rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With rs_kat
DelItem = Fg1.TextMatrix(Fg1.Row, 1)
If MsgBox("Вы хотите удалить " + Fg1.TextMatrix(Fg1.Row, 2) + " " + Fg1.TextMatrix(Fg1.Row, 3) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If rs_kat("Код") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
Fg1.DataRefresh
If .EOF Then .MoveLast
End If
End With

End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
rs_kat.UpdateBatch
Mconn.Execute ("UPDATE Kategor INNER JOIN Socmin ON Kategor.Код = Socmin.KodKategor SET Socmin.Kategor = [Kategor]![Name_Kategor]")
End Sub

Private Sub Fg1_Click()
If (Fg1.TextMatrix(0, Fg1.Col)) = "Период" Then

Cal.Calendar1.DataChanged = True
Cal.Calendar1.Value = Fg1.Cell(flexcpText)
Cal.Show
Socmin.Fg1.Cell(flexcpText) = MainForm.TMP1

rs_kat.UpdateBatch

'PeriodR
End If

If (Fg1.TextMatrix(0, Fg1.Col)) = "Код" Then
Cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
Cl = Cl + CStr(Combo_RS("Код")) & vbTab & Combo_RS("Name_Kategor") + "|"

Combo_RS.MoveNext
Loop
Fg1.ComboList = Cl
Else: Fg1.ComboList = ""
End If
End Sub

Private Sub FG1_FilterData(ByVal Row As Long, ByVal Col As Long, Value As String, ByVal SavingToDB As Boolean, WantThisCol As Boolean)
'Фильтр
 If Fg1.ColKey(Col) <> "Период" Then Exit Sub
WantThisCol = True


End Sub

Private Sub Fg1_GotFocus()
Fg1.ComboList = ""
End Sub


Private Sub FG1_RowColChange()
Fg1.Refresh
End Sub


Private Sub Form_Load()
Fg1.ComboList = ""
 Fg1.Editable = False
'Me.Caption = ""

' open connection
 '  Set mconn = New ADODB.Connection
'  mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
  
 
  
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
'Rs_kat.Open "Socmin"
rs_kat.Open "SELECT Socmin.Код, Socmin.Period, Socmin.KodKategor, Socmin.Kategor, Socmin.koli, Socmin.Value From Socmin ORDER BY Socmin.KodKategor, Socmin.koli"

Set Fg1.DataSource = rs_kat

Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
Combo_RS.Open "Kategor"

'FG1.ColComboList(2) = "..."  ' date picker popup


' правопреемник recordset в сетку
   
   Fg1.FocusRect = flexFocusSolid
    Fg1.Editable = True
    Fg1.DataMode = flexDMBound
    
    Fg1.AutoSearch = flexSearchFromCursor
    Fg1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Rs_kat.Close
'mconn.Close
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
