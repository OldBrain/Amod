VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Potok 
   Caption         =   "Ввод данных"
   ClientHeight    =   8472
   ClientLeft      =   60
   ClientTop       =   756
   ClientWidth     =   11856
   LinkTopic       =   "Form8"
   ScaleHeight     =   8472
   ScaleWidth      =   11856
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11856
      _ExtentX        =   20913
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
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   11415
      _cx             =   20135
      _cy             =   12515
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
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483635
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Potok.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   8
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
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
      Editable        =   1
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
      BackColorFrozen =   16711680
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5325
      Top             =   3990
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
            Picture         =   "Potok.frx":015B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Potok.frx":026D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Potok.frx":037F
            Key             =   "OOFL"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3480
      TabIndex        =   4
      Top             =   8160
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Номер счета в банке"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   8160
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ввод данных потоком. Это окно для быстрого ввода данных обязательных для выполнения расчета"
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   11535
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Новый 
         Caption         =   "Новый"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Удалить"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu Закрыть 
         Caption         =   "Закрыть"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Potok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim conn As ADODB.Connection
Dim rsVvod As ADODB.Recordset
Public rsDom As ADODB.Recordset
Dim rsKv As ADODB.Recordset
Public Num As String
Dim clDom As String, clKv As String
Dim It As String


Private Sub Form_Unload(Cancel As Integer)
Mconn.Execute ("DELETE MainOccupant.OLDNUM From MainOccupant WHERE (((MainOccupant.OLDNUM) Is Null))")
Mconn.Execute ("UPDATE KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom SET MainOccupant.DomTip = [KLS_PODR]![Tip], MainOccupant.Подразд = [KLS_PODR]![Подразделение]")

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            Новый_Click
        Case "Delete"
            Удалить_Click
        Case "OOFL1"
            Закрыть_Click
    End Select
End Sub


Private Sub Form_Load()

'Set conn = New ADODB.Connection
 ' conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'conn.Open "data/Kvartplata.mdb"
  
'VS.Sort = flexSortUseColSort



Set rsVvod = New ADODB.Recordset
Set rsDom = New ADODB.Recordset
Set rsKv = New ADODB.Recordset

rsVvod.Open ("SELECT  MainOccupant.Numer, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.KV, MainOccupant.NLODGERF, MainOccupant.COMSPACE, " + Chr(34) + "..." + Chr(34) + "AS Lg ,MainOccupant.BanKN  FROM MainOccupant"), Mconn, adOpenKeyset, adLockPessimistic
Set VS.DataSource = rsVvod

rsDom.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num From KLS_PODR ORDER BY KLS_PODR.КОД"), Mconn
clDom = ""
rsDom.MoveFirst
Do While Not rsDom.EOF
If Len(rsDom!Num) <> "" Then
clDom = clDom + "# " + Str(rsDom("код").Value) + ";" + rsDom("NAIM_KLS").Value + " Дом№ " + rsDom("Num").Value + "|"
End If
rsDom.MoveNext
Loop

rsKv.Open ("SELECT TipKv.Код, TipKv.Name_Kv FROM TipKv"), Mconn

clKv = ""
rsKv.MoveFirst
Do While Not rsKv.EOF
clKv = clKv + "#" + Str(rsKv("Код")) + ";" + rsKv("Name_Kv") + "|"
rsKv.MoveNext
Loop


VS.ComboSearch = flexCmbSearchLists
VS.ColComboList(6) = clDom
VS.ColComboList(8) = clKv
End Sub

Private Sub VS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'MsgBox VS.TextMatrix(Row, Col)
If VS.TextMatrix(Row, Col) = "" Then VS.TextMatrix(Row, Col) = "0"
If Col = 2 Then VS.TextMatrix(Row, 12) = Numer(VS.TextMatrix(Row, 1), MainForm.Jak, MainForm.Ray)
If VS.Col = VS.Cols - 2 Then VS.Col = 1
VS.Col = VS.Col + 1
SendKeys "{Enter}"

Label3.Caption = VS.TextMatrix(VS.Row, 12)
End Sub

Private Sub VS_Click()
Label3.Caption = VS.TextMatrix(VS.Row, 12)

End Sub

Private Sub VS_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)

'If VS.ComboIndex >= 0 Then
'If Col = 6 Then VS.TextMatrix(VS.Row, 6) = Val(VS.ComboItem())
'VS.ComboIndex = VS.ComboIndex + 1
'End If
'MsgBox FinishEdit

End Sub

Private Sub VS_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
'If It <> "" Then VS.ComboIndex = It


End Sub

Private Sub VS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If Col = 12 Then Cancel = True
If Col = 1 Then Cancel = True

If Col = 11 Then
Num = VS.TextMatrix(Row, 1)
'Potok.Hide
'Fmain.Show
End If

End Sub

Private Sub Закрыть_Click()
Me.Hide
Unload Me
End Sub

Private Sub Новый_Click()
VS.AddItem "х"
VS.Col = 2
VS.SetFocus
SendKeys "{Enter}"
VS.ColComboList(6) = clDom
VS.ColComboList(8) = clKv
End Sub

Private Sub Удалить_Click()
If MsgBox("Удалить " + VS.TextMatrix(VS.Row, 3) + " " + VS.TextMatrix(VS.Row, 4) + " " + VS.TextMatrix(VS.Row, 5) + " ?", vbYesNo) = vbYes Then
If MsgBox("ВНИМАНИЕ!! будут удалены все данные о начислениях и оплате по данному лицевому счету ", vbYesNo) = vbYes Then
VS.RemoveItem VS.Row
VS.Refresh
End If
End If
End Sub
