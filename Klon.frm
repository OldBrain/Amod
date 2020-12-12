VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Doma 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6948
   ClientLeft      =   120
   ClientTop       =   -276
   ClientWidth     =   12048
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H8000000A&
   Icon            =   "Klon.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1004
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Добавить F4"
      Height          =   315
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "Klon.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Удалить <F8>"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Удалить F8"
      Height          =   315
      Left            =   1560
      MousePointer    =   1  'Arrow
      Picture         =   "Klon.frx":0965
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Удалить <F8>"
      Top             =   6600
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11655
      _cx             =   20558
      _cy             =   10398
      Appearance      =   2
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
      BackColor       =   14342328
      ForeColor       =   -2147483640
      BackColorFixed  =   9922048
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   14342328
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   0
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
      FormatString    =   $"Klon.frx":0FC0
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
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
      CellButtonPicture=   "Klon.frx":1160
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
      Height          =   180
      Left            =   1560
      Picture         =   "Klon.frx":31A2
      ToolTipText     =   "Закрыть окно"
      Top             =   0
      Width           =   180
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Справочник домов"
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   11490
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1080
      Picture         =   "Klon.frx":35B0
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   0
      Picture         =   "Klon.frx":3CFA
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   2040
      Picture         =   "Klon.frx":4444
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "Doma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Combo_RS As ADODB.Recordset



Private Sub DataList1_Click()
DataList1.Refresh
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command12_Click()

Dim n, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
n = -10
Max = -10
rs_kat.MoveFirst

Do While Not rs_kat.EOF

If rs_kat("Код").Value = "" Then
rs_kat.Delete
rs_kat.Requery
rs_kat.MoveFirst
End If

N1 = rs_kat("Код").Value

'MsgBox Rs_kat("Код").Value

If N1 >= n Then
n = N1
Max = N1
End If

rs_kat.MoveNext
Loop

rs_kat.AddNew
rs_kat("Код") = n + 1
rs_kat("NAIM_KLS") = "Новый адрес"
rs_kat.UpdateBatch
fg1.DataRefresh
rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim RSdel As ADODB.Recordset
Set RSdel = New ADODB.Recordset

Dim DelItem As String
With rs_kat
DelItem = fg1.TextMatrix(fg1.Row, 1)
If MsgBox("Вы хотите удалить " + fg1.TextMatrix(fg1.Row, 2) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''

RSdel.Open ("SELECT MainOccupant.Dom From MainOccupant WHERE (((MainOccupant.Dom)=" + DelItem + "))"), Mconn, adOpenKeyset, adLockPessimistic

If RSdel.RecordCount > 0 Then
MsgBox "Удалять нельзя, есть жильцы по этому адресу."
Exit Sub
End If
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
rs_kat.UpdateBatch
Mconn.Execute ("UPDATE KLS_PODR INNER JOIN TipDom ON KLS_PODR.Tip = TipDom.Код SET KLS_PODR.Tip_Naim = [TipDom]![Name_Dom]")
'MsgBox (FG1.IsSelected(4))


End Sub
Private Sub FG1_RowColChange()
fg1.Refresh
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
      Unload Me
   End If

End Sub

Private Sub Form_Load()
Me.KeyPreview = True
MakeWindow Me, False
 
 
 
 fg1.Editable = False
' open connection
  ' Set mconn = New ADODB.Connection
'  mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
    
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
 
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
rs_kat.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.Подразделение, KLS_PODR.благ, KLS_PODR.Imp From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")
Set fg1.DataSource = rs_kat



Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
Combo_RS.Open "TipDom"



' правопреемник recordset в сетку
   
   fg1.FocusRect = 3
    'flexFocusSolid
    fg1.Editable = True
    fg1.DataMode = flexDMBound
        
    fg1.AutoSearch = flexSearchFromCursor
    fg1.ExplorerBar = flexExSortShowAndMove
    
    
    ' Cвойства, свойства необходимые для сортировки в этом гриде не работают
    ' из за строки поиска
    fg1.AllowUserResizing = flexResizeBoth
    fg1.ExtendLastCol = True
    
    

    
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

'DoEvents
Pod.Show
Pod.Label1.Caption = "Пожалуйста подождите, идет проверка адресов и тарифов"
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = 1000
Pod.ProgressBar1.Value = 100

Mconn.Execute ("UPDATE KLS_PODR SET KLS_PODR.Num = '0' WHERE (((KLS_PODR.Num) Is Null))")
Pod.ProgressBar1.Value = 300
Mconn.Execute ("UPDATE KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom SET MainOccupant.DomTip = [KLS_PODR]![Tip], MainOccupant.Подразд = [KLS_PODR]![Подразделение]")
Pod.ProgressBar1.Value = 700
Mconn.Execute ("UPDATE MainOccupant INNER JOIN Adding ON MainOccupant.Numer = Adding.KodKv SET Adding.TipDomKod = [MainOccupant]![DomTip]")
Pod.ProgressBar1.Value = 1000
rs_kat.UpdateBatch
rs_kat.Close

Unload Pod
MainMenu.Show

'Mconn.Close
End Sub

Private Sub Label1_Click()

End Sub



Private Sub Закрыть_Click()
Command1_Click
End Sub



Private Sub FG1_Click()

If (fg1.TextMatrix(0, fg1.Col)) = "Тип" Then
Cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
Cl = Cl + CStr(Combo_RS("Код")) & vbTab & Combo_RS("Name_Dom") + "|"

Combo_RS.MoveNext
Loop
fg1.ComboList = Cl

Else: fg1.ComboList = ""
End If

End Sub

Private Sub imgTitleHelp_Click()

Unload Doma
End Sub


