VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Constant 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3255
   ClientLeft      =   15
   ClientTop       =   75
   ClientWidth     =   4980
   ControlBox      =   0   'False
   Icon            =   "Constant.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Удалить<F8>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Добавить<F4>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть<F12>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      _cx             =   8493
      _cy             =   4048
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Constant.frx":0CCA
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1890
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
      Height          =   210
      Left            =   3960
      Picture         =   "Constant.frx":0D47
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   2040
      Picture         =   "Constant.frx":118F
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   2280
      Picture         =   "Constant.frx":18D9
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1800
      Picture         =   "Constant.frx":2023
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Constant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Addrconn As ADODB.Recordset
Dim Combo_RS As ADODB.Recordset
Dim F, f1, sq, sq1, SQDOM As String

Private Sub FG1_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
'FG1.TextMatrix(FG1.Row, 3) = Str(FG1.ComboCount)


End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Command1_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Command1_Click
'MsgBox (KeyAscii)
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





Private Sub Command1_Click()
Unload Me
Constant.Hide

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
Rs_kat("KodNach") = "0"
Rs_kat("NameNach") = " "
'Rs_kat("Im") = "Имя"
'Rs_kat("Ot") = "Отчество"
Rs_kat.UpdateBatch
FG1.DataRefresh
Rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With Rs_kat
DelItem = FG1.TextMatrix(FG1.Row, 1) + FG1.TextMatrix(FG1.Row, 2) + FG1.TextMatrix(FG1.Row, 3)
'MsgBox (DelItem)
If MsgBox("Вы хотите удалить начисление>" + FG1.TextMatrix(FG1.Row, 3) + "  код> " + FG1.TextMatrix(FG1.Row, 2) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If CStr(Rs_kat("Numer")) + CStr(Rs_kat("KodNach")) + Rs_kat("NameNach") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
FG1.DataRefresh
'If .EOF Then .MoveLast
End If
End With
End Sub


Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
         If Combo_RS("Kod") = FG1.TextMatrix(FG1.Row, 2) Then
    FG1.TextMatrix(FG1.Row, 3) = Combo_RS("Naim")
                      End If
Combo_RS.MoveNext
Loop






Rs_kat.UpdateBatch
'mconn.Execute ("UPDATE Constanta INNER JOIN Nachisleniy ON Constanta.KodNach = Nachisleniy.Kod SET Constanta.NameNach = [Nachisleniy]![Naim]")
End Sub

Private Sub FG1_Click()
If (FG1.TextMatrix(0, FG1.Col)) = "Код" Then
Cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
Cl = Cl + CStr(Combo_RS("Kod")) & vbTab & Combo_RS("Naim") + "|"
Combo_RS.MoveNext
Loop
FG1.ComboList = Cl
Else: FG1.ComboList = ""
End If
End Sub

Private Sub Form_Activate()
'Form_Load
End Sub

Private Sub Form_Load()

MakeWindow Me, True

    FG1.FocusRect = flexFocusSolid
    FG1.Editable = flexEDKbdMouse
    FG1.DataMode = flexDMBound
    FG1.AutoSearch = flexSearchFromCursor
    FG1.ExplorerBar = flexExSortShowAndMove
 
' open connection
   'Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"
 
 ' Рекордсет для сетки
Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = Mconn
'f = Filter.nm
F = Filter.Nm
If F <> "" Then f1 = "WHERE (((Constanta.Numer)=" & F & "))"
'MsgBox (f + "  111")
'sq = "SELECT OtheOwner.Numer, OtheOwner.Dom, OtheOwner.KV, OtheOwner.FAM, OtheOwner.IM, OtheOwner.OT, OtheOwner.PRIVILEGE, OtheOwner.BIRTHDAY, OtheOwner.NFAMILY, OtheOwner.PASSPORT, OtheOwner.LDATEBEG, OtheOwner.LDATEEND From OtheOwner WHERE (((OtheOwner.Numer)=" & f & "))"
sq = "SELECT Constanta.Numer, Constanta.KodNach, Constanta.NameNach  From Constanta " & f1
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

Combo_RS.Open "Nachisleniy"


' Привязываем Рекордсет к сетке грида
Set FG1.DataSource = Rs_kat
FG1.Refresh

'If Rs_kat.BOF = False Or Rs_kat.EOF = False Then
'Rs_kat.Fields.Item.Value

' Назначаем подписи вверху экрана
'Label1.Caption = Rs_kat.Fields("Numer").Value
'If AddrConn.EOF = False Then AddrConn.MoveFirst
'Label1.Caption = Rs_kat.Fields("Numer").Value
'If AddrConn.EOF = False Then Adres = "ул." + AddrConn.Fields("NAIM_KLS").Value + ", дом № " + AddrConn.Fields("Num").Value + " тип дома: " + CStr(AddrConn.Fields("Tip").Value) + " -> " + AddrConn.Fields("Tip_Naim").Value
'If Adres <> "" Then Label4.Caption = Adres
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs_kat.UpdateBatch
'AddrConn.Close
'Rs_kat.Close
'Combo_RS.Close
'MsgBox ("Закрытие")
'Rs_kat.Close
'mconn.Close
'Filter.Refresh
'Filter.FG.Redraw
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
