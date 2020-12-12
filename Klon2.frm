VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tarif 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Тарифы"
   ClientHeight    =   6912
   ClientLeft      =   156
   ClientTop       =   456
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6912
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Добавить новый тариф"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Удалить"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Сохранить и выйти"
            ImageKey        =   "Save"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      _cx             =   20981
      _cy             =   11245
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
      BackColor       =   -2147483626
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483626
      GridColor       =   -2147483630
      GridColorFixed  =   -2147483631
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Klon2.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   4
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
      ComboSearch     =   1
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
      Left            =   5415
      Top             =   2970
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
            Picture         =   "Klon2.frx":0152
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Klon2.frx":0264
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Klon2.frx":0376
            Key             =   "Save"
         EndProperty
      EndProperty
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
      Left            =   0
      TabIndex        =   3
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
      Height          =   156
      Left            =   5760
      Picture         =   "Klon2.frx":0488
      Top             =   0
      Width           =   156
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Т А Р И Ф Ы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11775
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      WindowList      =   -1  'True
      Begin VB.Menu Добавить 
         Caption         =   "Добавить"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Удалить"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Закрыть 
         Caption         =   "Закрыть <Ok>"
      End
   End
End
Attribute VB_Name = "Tarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kod_tar As String
Public Tar As Double
Public num_tar As String
Dim Cl As String
Dim rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection
'Dim Combo_Conn As ADODB.Connection
Dim Combo_RS As ADODB.Recordset
Dim comboKV As ADODB.Recordset
Dim comboDOM As ADODB.Recordset

Private Sub Fg1_DblClick()

'Если колонка равна 9/10 или 11 то мы настраиваем затраты по тарифу

             If fg1.Col = 9 Or fg1.Col = 10 Or fg1.Col = 11 Then

'Сначала определяем какой тариф настраиваем основной, доп или ....
If fg1.Col = 9 Then
num_tar = 1
End If
If fg1.Col = 10 Then
num_tar = 2
End If
If fg1.Col = 11 Then
num_tar = 3
End If
'теперь пробуем привязать тариф к затратам
Me.Tar = fg1.TextMatrix(fg1.Row, fg1.Col)
kod_tar = fg1.TextMatrix(fg1.Row, 1)
Me.Enabled = False

'теперь обращаемся к прогцедуре mainform.zatr_tarif для заполнения файла zatr_tarif
MainForm.zatr_tarif

'MsgBox FG1.TextMatrix(FG1.Row, 1)
zatr_tarif.Show
zatr_tarif.Label2.Caption = "Категория расчета " + fg1.TextMatrix(fg1.Row, 4) + " Тариф >" + fg1.TextMatrix(fg1.Row, fg1.Col)

            End If
End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim nr As Long, nc As Long      'при каждом движении мыши вычисляется № строки и столбца
    
    On Error GoTo ex
    Static R As Long, c As Long     'эти №№ изменяются при переходе границы ячейки
    nr = fg1.MouseRow:    nc = fg1.MouseCol  ' get coordinates
    
    If nr < 1 Or nc = -1 Then
    fg1.ToolTipText = ""
    Exit Sub
    End If
    If c <> nc Or R <> nr Then                   ' update tooltip text
        
       If fg1.TextMatrix(nr, nc) <> "" Then
        fg1.ToolTipText = fg1.TextMatrix(nr, nc)
        End If
        R = nr:            c = nc
        DoEvents
    End If
ex:

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "New"
           
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
rs_kat("KATEGOR") = "Новая запись"
rs_kat.UpdateBatch
fg1.DataRefresh
rs_kat.MoveLast
End If
        Case "Delete"
            Dim DelItem As String
With rs_kat
DelItem = fg1.TextMatrix(fg1.Row, 1)
If MsgBox("Вы хотите удалить " + fg1.TextMatrix(fg1.Row, 2) + " " + fg1.TextMatrix(fg1.Row, 3) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If rs_kat("Код") = DelItem Then .Delete

Mconn.Execute ("DELETE zatr_tarif.kod_tar, zatr_tarif.* From zatr_tarif WHERE (((zatr_tarif.kod_tar)=" + DelItem + "))")

If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
fg1.DataRefresh
If .EOF Then .MoveLast
End If
End With
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
Pod.ProgressBar1.Max = 1000
Pod.Show
Pod.Label1.Font = 8
Pod.Label1.Caption = " Пожалуйста подождите. Идет проверка и расстановка тарифов."
Pod.Refresh
Pod.Label1.Refresh

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
   Next
   
Tarif.Enabled = False
'Jdite.Show
'Jdite.Label1.Refresh
rs_kat.UpdateBatch

Mconn.Execute ("UPDATE KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom SET MainOccupant.DomTip = [KLS_PODR]![Tip], MainOccupant.Подразд = [KLS_PODR]![Подразделение]")
Mconn.Execute ("UPDATE MainOccupant INNER JOIN Adding ON MainOccupant.Numer = Adding.KodKv SET Adding.TipDomKod = [MainOccupant]![DomTip]")

For I = 251 To 500
    Pod.ProgressBar1.Value = I
   Next
'Mconn.Execute ("UPDATE Adding INNER JOIN Tarif ON (Adding.TipDomKod = Tarif.KodDOM) AND (Adding.TipKvKod = Tarif.KodKV) AND (Adding.KodKat = Tarif.KodKat) SET Adding.Tarif = [Tarif]![Value], Adding.TarifI = [Tarif]![TarifI], Adding.TarifD = [Tarif]![TarifD], Adding.kod_tar = [Tarif]![Код], Adding.norm = [Tarif]![norm]")

Mconn.Execute ("UPDATE Adding INNER JOIN Tarif ON (Adding.KodKat = Tarif.KodKat) AND (Adding.TipKvKod = Tarif.KodKV) AND (Adding.TipDomKod = Tarif.KodDOM) SET Adding.Tarif = [Tarif]![Value], Adding.TarifI = [Tarif]![TarifI], Adding.TarifD = [Tarif]![TarifD], Adding.kod_tar = [Tarif]![Код], Adding.norm = [Tarif]![norm]")



For I = 501 To 750
    Pod.ProgressBar1.Value = I
   Next

'Mconn.Execute ("UPDATE KLS_PODR INNER JOIN (Adding INNER JOIN Tarif ON (Adding.TipKvKod = Tarif.KodKV) AND (Adding.KodKat = Tarif.KodKat)) ON (KLS_PODR.Tip = Tarif.KodDOM) AND (KLS_PODR.КОД = Adding.TipDomKod) SET Adding.Tarif = [Tarif]![Value], Adding.TarifI = [Tarif]![TarifI], Adding.TarifD = [Tarif]![TarifD]")
Mconn.Execute ("UPDATE Adding SET Adding.Tarif = 0 WHERE ((([Adding].[Tarif]) Is Null))")
Mconn.Execute ("UPDATE Adding SET Adding.norm = 0 WHERE ((([Adding].[norm]) Is Null))")
For I = 751 To 1000
    Pod.ProgressBar1.Value = I
   Next



'Mconn.Execute ("UPDATE Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd SET tmp_lgota.tarif = [Adding]![Tarif]")

Unload Pod
Unload Me
MainMenu.Enabled = True
MainMenu.Show

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
rs_kat("KATEGOR") = "Новая запись"
rs_kat.UpdateBatch
fg1.DataRefresh
rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With rs_kat
DelItem = fg1.TextMatrix(fg1.Row, 1)
If MsgBox("Вы хотите удалить " + fg1.TextMatrix(fg1.Row, 2) + " " + fg1.TextMatrix(fg1.Row, 3) + "?", vbYesNo) = vbYes Then
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
Dim Er As Label
' Tarif.FG1.Cell(flexcpText) = MainForm.TMP1
On Error GoTo Er
rs_kat.UpdateBatch
Mconn.Execute ("UPDATE Kategor INNER JOIN Tarif ON Kategor.Код = Tarif.KodKat SET Tarif.Kategor = [Kategor]![Name_Kategor]")
Mconn.Execute ("UPDATE TipKV INNER JOIN Tarif ON TipKV.Код = Tarif.KodKV SET Tarif.NameKV = [TipKV]![Name_KV]")
Mconn.Execute ("UPDATE TipDOM INNER JOIN Tarif ON TipDOM.Код = Tarif.KodDOM SET Tarif.NameDOM = [TipDOM]![Name_DOM]")
fg1.ComboList = ""
Exit Sub
Er: MsgBox ("Повторите ввод")
End Sub






Private Sub FG1_Click()

If (fg1.TextMatrix(0, fg1.Col)) = "Период" Then
Cal1.Calendar1.DataChanged = True
Cal1.Calendar1.Value = fg1.Cell(flexcpText)
Cal1.Show
rs_kat.UpdateBatch
End If

If (fg1.TextMatrix(0, fg1.Col)) = "Код" Then
Cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
Cl = Cl + CStr(Combo_RS("Код")) & vbTab & Combo_RS("Name_Kategor") + "|"
Combo_RS.MoveNext
Loop
fg1.ComboList = Cl
'Else: FG1.ComboList = ""
End If

If (fg1.TextMatrix(0, fg1.Col)) = "ТипКВ" Then
Cl = ""
comboKV.MoveFirst
Do While Not comboKV.EOF
Cl = Cl + CStr(comboKV("Код")) & vbTab & comboKV("Name_Kv") + "|"
comboKV.MoveNext
Loop
fg1.ComboList = Cl
'Else: FG1.ComboList = ""
End If

If (fg1.TextMatrix(0, fg1.Col)) = "ТипДом" Then
Cl = ""
comboDOM.MoveFirst
Do While Not comboDOM.EOF
Cl = Cl + CStr(comboDOM("Код")) & vbTab & comboDOM("Name_DOM") + "|"
comboDOM.MoveNext
Loop
fg1.ComboList = Cl
'Else: FG1.ComboList = ""
End If
If (fg1.TextMatrix(0, fg1.Col)) = "Код" And (fg1.TextMatrix(0, fg1.Col)) = "ТипКВ" And (fg1.TextMatrix(0, fg1.Col)) = "ТипДом" Then fg1.ComboList = ""


End Sub

Private Sub Fg1_GotFocus()
'FG1.ComboList = ""

End Sub


Private Sub FG1_RowColChange()
fg1.Refresh
End Sub



Private Sub Form_Load()
fg1.ComboList = ""
 fg1.Editable = False

' open connection
 '  Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
  
 
  
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockPessimistic

'Rs_kat.Open "Tarif"

rs_kat.Open "SELECT Tarif.Код, Tarif.Period, Tarif.KodKat, Tarif.Kategor, Tarif.KodKV, Tarif.NameKV, Tarif.KodDOM, Tarif.NameDOM, Tarif.Value, Tarif.TarifI, Tarif.TarifD, Tarif.norm From Tarif ORDER BY Tarif.KodKat,  Tarif.KodDOM,Tarif.KodKV"
Set fg1.DataSource = rs_kat



Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
Combo_RS.Open "Kategor"

Set comboKV = New ADODB.Recordset
Set comboKV.ActiveConnection = Mconn
comboKV.Open "TipKV"

Set comboDOM = New ADODB.Recordset
Set comboDOM.ActiveConnection = Mconn
comboDOM.Open "TipDom"



'Combo_RS.CursorType = adOpenForwardOnly
'Combo_RS.LockType = adLockBatchOptimistic

'FG1.ColComboList(2) = "..."  ' date picker popup


' правопреемник recordset в сетку
   
   fg1.FocusRect = flexFocusSolid
    fg1.Editable = True
    fg1.DataMode = flexDMBound
    
   fg1.AutoSearch = flexSearchFromCursor
   fg1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Rs_kat.Close
'comboDOM.Close
'comboKV.Close
'Combo_RS.Close
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
