VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Nachisleniy 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Справочник начислений"
   ClientHeight    =   6672
   ClientLeft      =   156
   ClientTop       =   456
   ClientWidth     =   14100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6672
   ScaleWidth      =   14100
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14100
      _ExtentX        =   24871
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
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
      EndProperty
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Продолжить перенос данных"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   8175
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Удалить <F8>"
      Enabled         =   0   'False
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
      Left            =   3000
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Добавить<F4>"
      Enabled         =   0   'False
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
      Left            =   1560
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ok"
      Enabled         =   0   'False
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
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   15135
      _cx             =   26696
      _cy             =   10821
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
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483638
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483625
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Nachisleniy.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   2
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
      Left            =   6450
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
            Picture         =   "Nachisleniy.frx":01FB
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Nachisleniy.frx":030D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Nachisleniy.frx":041F
            Key             =   "Save"
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
Attribute VB_Name = "Nachisleniy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cl As String
Public FI As String
Public FIBl As String
Dim rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection
'Dim Combo_Conn As ADODB.Connection
Dim Combo_RS, Combo_rs1, Proverka As ADODB.Recordset
Public Old

Private Sub Command4_Click()
rs_kat.UpdateBatch
Unload Me
Mass.Show
Mass.продолжить
End Sub

Private Sub FG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Old = fg1.TextMatrix(fg1.Row, 5)
'FORM1 = Old
End Sub

Private Sub FG1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If fg1.Col = 1 Or fg1.Col = 3 Then
Cancel = True
End If
'Trim((FG1.TextMatrix(0, FG1.Col))) = "Формула" Or

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
'    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            Command2_Click
        Case "Delete"
            Command3_Click
        Case "Save"
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

For rw = 1 To fg1.Rows - 1
If fg1.TextMatrix(rw, 10) = "" Then
MsgBox "Заполните значения формул"
Exit Sub
End If
Next

Unload Sprav
Nachisleniy.Enabled = False
Jdite.Show
Jdite.Label1.Refresh
rs_kat.UpdateBatch
' Если не заполнены поля спр.начислений то проставляем Формула=0 Вид льготы="Не определено"
Mconn.Execute ("UPDATE Nachisleniy SET Nachisleniy.Vid = " + Chr(34) + "Не определено" + Chr(34) + " WHERE (((Nachisleniy.Vid) Is Null))")
Mconn.Execute ("UPDATE Nachisleniy SET Nachisleniy.formula = " + Chr(34) + "0" + Chr(34) + " WHERE (((Nachisleniy.formula) Is Null))")
Mconn.Execute ("UPDATE Nachisleniy SET Nachisleniy.SchetZ = " + Chr(34) + "Не определен" + Chr(34) + " WHERE (((Nachisleniy.SchetZ) Is Null))")
Mconn.Execute ("UPDATE nachisleniy SET nachisleniy.NDS = 0 WHERE (((nachisleniy.NDS) Is Null))")
Mconn.Execute ("UPDATE nachisleniy SET nachisleniy.Komis = 0 WHERE (((nachisleniy.Komis) Is Null))")

' Теперь обновляем Adding
Mconn.Execute ("UPDATE Nachisleniy INNER JOIN Adding ON Nachisleniy.Kod = Adding.KodN SET Adding.KodKat = [Nachisleniy]![КодKategor], Adding.NameKat = [Nachisleniy]![Kategor], Adding.Formula = [Nachisleniy]![Formula], Adding.Tip = [Nachisleniy]![Tip], Adding.Lig = [Nachisleniy]![Lig], Adding.LgotaVid = [Nachisleniy]![Vid], Adding.NameN = [Nachisleniy]![Naim], Adding.SchetZ = [Nachisleniy]![SchetZ], Adding.FormulaB = [Nachisleniy]![FormulaB], Adding.Sch = [nachisleniy]![Sch], Adding.edizm = [nachisleniy]![edizm]")

'mconn.Execute ("UPDATE Adding SET Adding.FormulaB = [Nachisleniy]![Formula] WHERE (((Adding.FormulaB) Is Null))")

Unload Nachisleniy
Unload Jdite
'Sprav.Show
'Sprav.Refresh
Sprav.Enabled = True
Sprav.Show
End Sub

Private Sub Command2_Click()

Dim n, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
n = 0
rs_kat.MoveFirst
Do While Not rs_kat.EOF
If rs_kat("Kod").Value = "" Then
rs_kat.Delete
rs_kat.MoveFirst
End If
N1 = rs_kat("Kod").Value
If N1 > n Then n = N1
rs_kat.MoveNext
Loop

rs_kat.AddNew
rs_kat("Kod") = n + 1
rs_kat("Naim") = "Новая вид расчета"
rs_kat.UpdateBatch
fg1.DataRefresh
rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
Dim Выход As Label
'Dim A2 As doumle
A2 = ""
Set Proverka = New ADODB.Recordset
Set Proverka.ActiveConnection = Mconn
Proverka.CursorType = adOpenForwardOnly
Proverka.LockType = adLockBatchOptimistic
Proverka.Open "SELECT Adding.KodN From Adding WHERE (((Adding.KodN)=" + fg1.TextMatrix(fg1.Row, 1) + "))"
On Error Resume Next
A2 = Proverka.Fields("KodN").Value
If A2 <> "" Then
MsgBox ("Удалять нельзя! используется в расчетах!")
Exit Sub
End If


With rs_kat
DelItem = fg1.TextMatrix(fg1.Row, 1)
If MsgBox("Вы хотите удалить N " + fg1.TextMatrix(fg1.Row, 1) + " " + fg1.TextMatrix(fg1.Row, 4) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If rs_kat("Kod") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
fg1.DataRefresh
If .EOF Then .MoveLast
End If
End With

End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Проверка
Mconn.Execute ("UPDATE Nachisleniy INNER JOIN Kategor ON Nachisleniy.КодKategor = Kategor.Код SET Nachisleniy.Kategor = [Kategor]![Name_Kategor]")
End Sub

Private Sub FG1_Click()
If Trim((fg1.TextMatrix(0, fg1.Col))) = "Форм.без льгот" Then
FI = fg1.TextMatrix(fg1.Row, fg1.Col)
Formula.Text1 = fg1.TextMatrix(fg1.Row, 5)
Formula.Text2 = fg1.TextMatrix(fg1.Row, 10)
Formula.Show 1

'Nachisleniy.Enabled = False
End If

If (fg1.TextMatrix(0, fg1.Col)) <> "Код" Or fg1.Col <> 9 Then fg1.ComboList = ""
If (fg1.TextMatrix(0, fg1.Col)) = "Код" Then
Cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF

Cl = Cl + CStr(Combo_RS("Код")) & vbTab & Combo_RS("Name_Kategor") + "|"
Combo_RS.MoveNext
Loop
fg1.ComboList = Cl

End If


If fg1.Col = 9 Then
Cl = ""
Combo_rs1.MoveFirst
Do While Not Combo_rs1.EOF

Cl = Cl + CStr(Combo_rs1("Schet_Name")) & vbTab & Combo_rs1("Schet") + "|"
Combo_rs1.MoveNext
Loop
fg1.ComboList = Cl

End If

End Sub

Private Sub Fg1_GotFocus()
'FG1.ComboList = ""
End Sub


Private Sub FG1_RowColChange()
'FG1.Refresh
End Sub


Private Sub Form_Load()
fg1.ComboList = ""

 
  
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
rs_kat.Open "Nachisleniy"




'FG1.RowHeight(0) = 1200
fg1.WordWrap = True
fg1.Cell(flexcpAlignment, 0, 0, 0, fg1.Cols - 1) = flexAlignCenterCenter



fg1.CellAlignment = flexAlignGeneralCenter




Set fg1.DataSource = rs_kat


fg1.RowHeight(0) = 900
fg1.WordWrap = True
fg1.Cell(flexcpAlignment, 0, 0, 0, fg1.Cols - 1) = flexAlignCenterCenter




Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
Combo_RS.Open "Kategor"


Set Combo_rs1 = New ADODB.Recordset
Set Combo_rs1.ActiveConnection = Mconn
Combo_rs1.CursorType = adOpenForwardOnly
Combo_rs1.LockType = adLockBatchOptimistic
Combo_rs1.Open "Schet"


' правопреемник recordset в сетку
   
   'FG1.FocusRect = flexFocusSolid
    fg1.Editable = flexEDKbdMouse
    fg1.DataMode = flexDMBoundImmediate
    
    
    
    'FG1.AutoSearch = flexSearchFromCursor
    'FG1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)
Sprav.Enabled = True
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
Private Sub Проверка()

Dim Er As Label
'If InStr(1, FG1.TextMatrix(FG1.Row, 5), "LgotaP") And FG1.TextMatrix(FG1.Row, 8) = "Нет" Then MsgBox ("Предупреждение! В формуле присутствует параметр <LgotaР>, хотя <уменьшение лготой> = <Нет>. Расчет будет производится по формуле указанной Вами, НО ОТЧЕТЫ ПО ЛЬГОТАМ СКОРЕЕ ВСЕГО БУДЕТ СОБРАНЫ НЕ ВЕРНО!")
'If InStr(1, FG1.TextMatrix(FG1.Row, 5), "LgotaP") = False And FG1.TextMatrix(FG1.Row, 8) = "Да" Then MsgBox ("Предупреждение! В формуле отсутствует параметр <LgotaР>, хотя <уменьшение лготой> = <Да>. Расчет будет производится по формуле указанной Вами, НО ОТЧЕТЫ ПО ЛЬГОТАМ СКОРЕЕ ВСЕГО БУДЕТ СОБРАНЫ НЕ ВЕРНО!")
'Form = FG1.TextMatrix(FG1.Row, 5)
'FORM1 = Formula.Fr
If InStr(1, FI, "LgotaP") And fg1.TextMatrix(fg1.Row, 8) = "Нет" Then MsgBox ("Предупреждение! В формуле присутствует параметр <LgotaР>, хотя <уменьшение лготой> = <Нет>. Расчет будет производится по формуле указанной Вами, НО ОТЧЕТЫ ПО ЛЬГОТАМ СКОРЕЕ ВСЕГО БУДЕТ СОБРАНЫ НЕ ВЕРНО!")
If InStr(1, FI, "LgotaP") = False And fg1.TextMatrix(fg1.Row, 8) = "Да" Then MsgBox ("Предупреждение! В формуле отсутствует параметр <LgotaР>, хотя <уменьшение лготой> = <Да>. Расчет будет производится по формуле указанной Вами, НО ОТЧЕТЫ ПО ЛЬГОТАМ СКОРЕЕ ВСЕГО БУДЕТ СОБРАНЫ НЕ ВЕРНО!")
'FORM1 = Formula.Fr

'On Error GoTo Er
'mconn.Execute ("UPDATE Adding_Err SET Adding_Err.SummaI = " + Trim(FORM1))
'FG1.Editable = flexEDKbdMouse
'MsgBox (FORM1)
'FG1.TextMatrix(FG1.Row, 5) = FORM1
'Exit Sub
'Er:
'MsgBox ("Ошибка в формуле" + Err.Description)
'FG1.TextMatrix(FG1.Row, 5) = Old
'Formula.Show
'Formula.Text1 = FORM1
'Nachisleniy.Enabled = False
End Sub
