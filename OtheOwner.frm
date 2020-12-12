VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form OtheOwner 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3390
   ClientLeft      =   15
   ClientTop       =   75
   ClientWidth     =   7905
   ControlBox      =   0   'False
   Icon            =   "OtheOwner.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Удалить<F8>"
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
      Left            =   2760
      Picture         =   "OtheOwner.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   240
      Picture         =   "OtheOwner.frx":0388
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Закрыть<F12>"
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
      Left            =   5280
      Picture         =   "OtheOwner.frx":0506
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _cx             =   13996
      _cy             =   3625
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
      Rows            =   20
      Cols            =   50
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"OtheOwner.frx":066E
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      Height          =   240
      Left            =   0
      Picture         =   "OtheOwner.frx":0AA7
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "АРМ ""Квартплата + "" Отчеты и Анализ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   6810
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   2040
      Picture         =   "OtheOwner.frx":0FE9
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   360
      Picture         =   "OtheOwner.frx":1733
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   960
      Picture         =   "OtheOwner.frx":1E7D
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "OtheOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Addrconn As ADODB.Recordset
Dim Combo_RS As ADODB.Recordset
Dim F, f1, sq, sq1, SQDOM As String
Public Othe1 As Long

Private Sub FG1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub









Private Sub Command1_Click()
Othe1 = 0
'OtheOwner.Hide
Unload Me
'Form_Unload (0)
End Sub

Private Sub Command2_Click()
Rs_kat.UpdateBatch
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then

If Rs_kat.EOF = False Then Rs_kat.MoveFirst

Do While Not Rs_kat.EOF
If Rs_kat("Numer").Value = "" Then
Rs_kat.Delete
Rs_kat.MoveFirst
End If
Rs_kat.MoveNext
Loop

Rs_kat.AddNew
Rs_kat("Numer") = F
Rs_kat("Fam") = "Фамилия"
Rs_kat("PRIVILEGE") = 0
Rs_kat("Im") = "Имя"
Rs_kat("Ot") = "Отчество"

Rs_kat.UpdateBatch
FG1.DataRefresh
Rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim Prov As ADODB.Recordset
Set Prov = New ADODB.Recordset


Dim DelItem As String
With Rs_kat
DelItem = FG1.TextMatrix(FG1.Row, 13)

Prov.Open ("SELECT tmp_lgota.KodKv, tmp_lgota.OtheCode From tmp_lgota WHERE (((tmp_lgota.OtheCode)=" + DelItem + "))"), Mconn, adOpenKeyset, adLockPessimistic

If Prov.RecordCount <> 0 Then
If MsgBox("У " + FG1.TextMatrix(FG1.Row, 4) + " " + FG1.TextMatrix(FG1.Row, 5) + " есть льгота, удалить?", vbYesNo) = vbNo Then
Exit Sub
Else
Mconn.Execute ("DELETE tmp_lgota.OtheCode From tmp_lgota WHERE (((tmp_lgota.OtheCode)=" + DelItem + "))")
Mconn.Execute ("DELETE Lgota.OhteCode From Lgota WHERE (((Lgota.OhteCode)=" + DelItem + "))")
End If
End If

'MsgBox (DelItem)
If MsgBox("Вы хотите удалить " + DelItem + " " + FG1.TextMatrix(FG1.Row, 4) + " " + FG1.TextMatrix(FG1.Row, 5) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
'fg1.Delete
.MoveFirst
Do While Not .EOF
If Rs_kat("OhteCode") = DelItem Then
.Delete
Mconn.Execute ("DELETE tmp_lgota.OtheCode From tmp_lgota WHERE (((tmp_lgota.OtheCode)=" + DelItem + "))")
End If
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
FG1.DataRefresh
'If .EOF Then .MoveLast


End If
End With
End Sub


Private Sub FG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Rs_kat.UpdateBatch
End Sub

Private Sub FG1_Click()
'Dim othe As Double
If (FG1.TextMatrix(0, FG1.Col)) = "Льгота" Then
'cl = ""
'Combo_RS.MoveFirst
'Do While Not Combo_RS.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
'cl = cl + CStr(Combo_RS("N_KLS")) & vbTab & Combo_RS("Name_Kls") + "|"

'Combo_RS.MoveNext
'Loop
'fg1.ComboList = cl
'Else: fg1.ComboList = ""
Othe1 = FG1.TextMatrix(FG1.Row, 13)
DropForm2.Show
DropForm3.Show
DropForm3.Move DropForm2.Width + 1, (DropForm2.Height - DropForm3.Height) / 2
End If
End Sub

Private Sub Form_Activate()
'Form_Load
End Sub

Private Sub Form_Load()
MakeWindow Me, True

Filter.Enabled = False
Othe1 = 0
    FG1.FocusRect = flexFocusSolid
    FG1.Editable = flexEDKbdMouse
    FG1.DataMode = flexDMBound
    FG1.AutoSearch = flexSearchFromCursor
    FG1.ExplorerBar = flexExSortShowAndMove
 
' open connection
 '  Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
 
 ' Рекордсет для сетки
Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = Mconn
'f = Filter.nm
F = Filter.Nm
If F <> "" Then f1 = "WHERE (((OtheOwner.Numer)=" & F & "))"
'MsgBox (f + "  111")
'sq = "SELECT OtheOwner.Numer, OtheOwner.Dom, OtheOwner.KV, OtheOwner.FAM, OtheOwner.IM, OtheOwner.OT, OtheOwner.PRIVILEGE, OtheOwner.BIRTHDAY, OtheOwner.NFAMILY, OtheOwner.PASSPORT, OtheOwner.LDATEBEG, OtheOwner.LDATEEND From OtheOwner WHERE (((OtheOwner.Numer)=" & f & "))"
sq = "SELECT OtheOwner.Numer, OtheOwner.Dom, OtheOwner.KV, OtheOwner.FAM, OtheOwner.IM, OtheOwner.OT, OtheOwner.PRIVILEGE, OtheOwner.BIRTHDAY, OtheOwner.NFAMILY, OtheOwner.PASSPORT, OtheOwner.LDATEBEG, OtheOwner.LDATEEND, OtheOwner.OhteCode From OtheOwner " & f1
' Рекордсет для подписи (адреса)
Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
SQDOM = "SELECT OtheOwner.Numer, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim FROM OtheOwner LEFT JOIN KLS_PODR ON OtheOwner.Dom = KLS_PODR.КОД"

Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic

' Рекордсет для падающего списка лгот
Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
'MsgBox (sq)
'Открываем Рекордсеты
If F <> "" And F <> "Номер" Then Rs_kat.Open (sq) Else OtheOwner.Hide


Addrconn.Open (SQDOM)

Combo_RS.Open "KLS_PRIV"


' Привязываем Рекордсет к сетке грида
Set FG1.DataSource = Rs_kat
FG1.Refresh

'If Rs_kat.BOF = False Or Rs_kat.EOF = False Then
'Rs_kat.Fields.Item.Value

' Назначаем подписи вверху экрана
'Label1.Caption = Rs_kat.Fields("Numer").Value
If Addrconn.EOF = False Then Addrconn.MoveFirst
'Label1.Caption = Rs_kat.Fields("Numer").Value
adres = " "
'If AddrConn.EOF = False Then adres = "ул." + AddrConn.Fields("NAIM_KLS").Value + ", дом № " + AddrConn.Fields("Num").Value + " тип дома: " + CStr(AddrConn.Fields("Tip").Value) + " -> " + AddrConn.Fields("Tip_Naim").Value
'Label4.Caption = adres
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs_kat.UpdateBatch
Filter.Enabled = True
End Sub



Private Sub Выход_Click(Index As Integer)
Command1_Click

End Sub

Private Sub Добавить_Click(Index As Integer)
Command2_Click
End Sub

Private Sub Удалить_Click(Index As Integer)
Command3_Click
End Sub

Private Sub imgTitleHelp_Click()
Command1_Click
End Sub
