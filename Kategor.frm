VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Kategor 
   BackColor       =   &H80000001&
   Caption         =   "Справочник категорий расчетов"
   ClientHeight    =   6672
   ClientLeft      =   168
   ClientTop       =   468
   ClientWidth     =   7032
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form7"
   ScaleHeight     =   6672
   ScaleWidth      =   7032
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7032
      _ExtentX        =   12404
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "imlToolbarIcons(2)"
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
      BackColor       =   &H00C0FFFF&
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
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
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
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   495
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   6132
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7332
      _cx             =   12933
      _cy             =   10816
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Kategor.frx":0000
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
      Ellipsis        =   0
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
      Index           =   0
      Left            =   2055
      Top             =   3090
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
            Picture         =   "Kategor.frx":0124
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kategor.frx":0236
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kategor.frx":0348
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   2910
      Top             =   3090
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
            Picture         =   "Kategor.frx":045A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kategor.frx":056C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kategor.frx":067E
            Key             =   "OOFL"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   2
      Left            =   2910
      Top             =   3090
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
            Picture         =   "Kategor.frx":0998
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kategor.frx":0AAA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kategor.frx":0BBC
            Key             =   "OOFL"
         EndProperty
      EndProperty
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
Attribute VB_Name = "Kategor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_kat As ADODB.Recordset
Dim Comb As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Kr As String


Private Sub FG1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If fg1.Col = 1 Then Cancel = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            Command2_Click
        Case "Delete"
    Command3_Click
        Case "OOFL1"
            Command1_Click
    End Select
End Sub




Private Sub Command4_Click()
'PrintW.Show
'FG1.PrintGrid ("111")
'PrintSelection fg1
'fg1.PrintGrid "My Grid"
PrintW.Show
        PrintW.VP.StartDoc
        PrintW.VP.RenderControl = fg1.hwnd
        PrintW.VP.EndDoc
End Sub




Private Sub DataList1_Click()
DataList1.Refresh
End Sub

Private Sub Command1_Click()
rs_kat.UpdateBatch
' Обновляем Adding
Mconn.Execute ("UPDATE Adding INNER JOIN Kategor ON Adding.KodKat = Kategor.Код SET Adding.NameKat = [Kategor]![Name_Kategor], Adding.Parametr = [Kategor]![Parametr]")
' Обновляем спр. Начислений
Mconn.Execute ("UPDATE Kategor INNER JOIN Nachisleniy ON Kategor.Код = Nachisleniy.КодKategor SET Nachisleniy.Kategor = [Kategor]![Name_Kategor]")
Kategor.Hide
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
rs_kat("NAME_KATEGOR") = "Новая запись"
rs_kat.UpdateBatch
fg1.DataRefresh
rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With rs_kat
DelItem = fg1.TextMatrix(fg1.Row, 1)
If MsgBox("Вы хотите удалить " + fg1.TextMatrix(fg1.Row, 2) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If rs_kat("Код") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
fg1.DataRefresh
If .EOF Then .MoveLast
End If
End With

End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'FG1.TextMatrix(fg1.row, FG1.Col)=val(cl

rs_kat.UpdateBatch
End Sub

Private Sub FG1_Click()
'MsgBox (FG1.Cell(flexcpText))
'MsgBox (FG1.TextMatrix(FG1.Row, 1))
Kr = "1"
Kr = Trim(Str(fg1.TextMatrix(fg1.Row, 1)))
If fg1.TextMatrix(0, fg1.Col) = "Код опл" Then
Comb.Open ("SELECT nachisleniy.Kod, nachisleniy.КодKategor, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy WHERE (((nachisleniy.КодKategor)=" + Kr + ") AND ((nachisleniy.Tip)=" + Chr(34) + "-" + Chr(34) + "))")
Cl = ""
On Error Resume Next
Comb.MoveFirst
Do While Not Comb.EOF
If Comb.Fields("naim").Value <> "" Then Cl = Cl + CStr(Comb.Fields("Kod").Value) & vbTab & Comb.Fields("naim").Value + "|"

Comb.MoveNext
Loop
fg1.ComboList = Cl
Comb.Close
Else
fg1.ComboList = ""
End If

End Sub

Private Sub FG1_RowColChange()
fg1.Refresh

End Sub

Private Sub Form_Load()

 fg1.Editable = False

' open connection
  ' Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"
    
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
 
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
rs_kat.Open "Kategor"
Set fg1.DataSource = rs_kat

Set Comb = New ADODB.Recordset
Set Comb.ActiveConnection = Mconn
 
Comb.CursorType = adOpenStatic

'adOpenForwardOnly
Comb.LockType = adLockBatchOptimistic

'Comb.Open ""






' правопреемник recordset в сетку
   
   fg1.FocusRect = 3
    'flexFocusSolid
    fg1.Editable = True
    fg1.DataMode = flexDMBound
    
    fg1.AutoSearch = flexSearchFromCursor
    fg1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs_kat.Close
Mconn.Close
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
Private Sub PrintSelection(fg1 As VSFlexGrid, Row1&, Col1&, Row2&, Col2&)

    
        ' save current settings
        Dim hl%, tr&, LC&, rd%

        hl = fg1.HighLight: tr = fg1.TopRow: LC = fg1.LeftCol: rd = fg1.Redraw

        fg1.HighLight = 0
        fg1.Redraw = flexRDNone

    
        ' hide non-selected rows and columns
        Dim i&
    For i = fg1.FixedRows To fg1.Rows - 1

       If i < Row1 Or i > Row2 Then Fg.RowHidden(i) = True

    Next
    For i = fg1.FixedCols To fg1.Cols - 1

      If i < Col1 Or i > Col2 Then fg1.ColHidden(i) = True

    Next

    

    ' scroll to top left corner

    fg1.TopRow = fg1.FixedRows

    fg1.LeftCol = fg1.FixedCols

    

    ' print visible area

    fg1.PrintGrid

    

    ' restore control
    fg1.RowHidden(-1) = False

    fg1.ColHidden(-1) = False
    fg1.TopRow = tr: fg1.LeftCol = LC: fg1.HighLight = hl
    fg1.Redraw = rd
  End Sub


Sub PrintFlexGridOnVSPrinter()
        VP1.StartDoc
        VP1.RenderControl = fg1.hwnd
        
        VP1.EndDoc
    End Sub

