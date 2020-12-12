VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ReestrTablDoc 
   Caption         =   "Реестр документов"
   ClientHeight    =   7668
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   11244
   LinkTopic       =   "Form8"
   ScaleHeight     =   7668
   ScaleWidth      =   11244
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   552
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11244
      _ExtentX        =   19833
      _ExtentY        =   974
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "5"
                  Text            =   "Расчет"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Macro"
            Object.ToolTipText     =   "Macro"
            ImageKey        =   "Macro"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OOFL1"
            ImageKey        =   "OOFL"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid VStbd 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11055
      _cx             =   19500
      _cy             =   12303
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ReestrTablDoc.frx":0000
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5025
      Top             =   3585
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrTablDoc.frx":0176
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrTablDoc.frx":0288
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrTablDoc.frx":039A
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrTablDoc.frx":04AC
            Key             =   "OOFL"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Расчет 
         Caption         =   "Расчет"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Новый 
         Caption         =   "Новый документ"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Удалить"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Выход 
         Caption         =   "Выход"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "ReestrTablDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_kat As ADODB.Recordset
Public TM As ADODB.Recordset
Public Dl As ADODB.Recordset
Public n
'Dim mconn As ADODB.Connection

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Show
MainMenu.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
'    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            'ToDo: Add 'New' button code.
            Новый_Click
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            Удалить_Click
        Case "Macro"
            'ToDo: Add 'Macro' button code.
            MsgBox "Add 'Macro' button code."
        Case "OOFL1"
            'ToDo: Add 'OOFL1' button code.
            Выход_Click
    End Select
End Sub

Private Sub Form_Load()

'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
Set TM = New ADODB.Recordset
Set TM.ActiveConnection = Mconn
rs_kat.CursorType = adOpenKeyset
rs_kat.LockType = adLockOptimistic
rs_kat.Open ("ReestrTablDoc")
Set VStbd.DataSource = rs_kat


End Sub

Private Sub VStbd_DblClick()
DocTBL.Show
End Sub

Private Sub Выход_Click()

Unload Me
End Sub

Private Sub Новый_Click()
'НовыйТБЛ
ShapkaTBL.Show
End Sub
Public Sub НовыйТБЛ()
Dim N1 As Integer

If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
n = 0

'MsgBox (rs_kat.RecordCount)
'If Not Rs_kat.EOF = True Then
If Not rs_kat.EOF Then rs_kat.MoveFirst
Do While Not rs_kat.EOF
If rs_kat("Cod").Value = "" Then
rs_kat.Delete
rs_kat.MoveFirst
End If
N1 = rs_kat("Cod").Value
If N1 > n Then n = N1
rs_kat.MoveNext
Loop

rs_kat.AddNew
'Rs_kat("") = N + 1
rs_kat("Coment") = ShapkaTBL.Text1.Text
rs_kat("Nach") = ShapkaTBL.Combo2
rs_kat("Status") = 0
If Prostoy = True Then rs_kat("tip") = "Prostoy"
'If Odn = True Then Rs_kat("tip") = "ODN"

'?????????????????????????????
'Rs_kat("Tip") = DocShapka.Combo1.Text
 rs_kat("Data") = MainForm.DR
'DocShapka.Text1

rs_kat("NachCod") = ShapkaTBL.KNac
rs_kat("KodDom") = ShapkaTBL.KAdres
rs_kat("Adres") = ShapkaTBL.Combo3.Text






rs_kat.UpdateBatch
VStbd.DataRefresh
If Not rs_kat.EOF Then rs_kat.MoveLast
End If

End Sub

Private Sub Удалить_Click()
Set Dl = New ADODB.Recordset
Set Dl.ActiveConnection = Mconn
Dl.Open ("SELECT TablDoc.TablKod, TablDoc.Cod, TablDoc.TabNum From TablDoc WHERE (((TablDoc.Cod)=" + VStbd.TextMatrix(VStbd.Row, 1) + "))"), Mconn, adOpenKeyset, adLockPessimistic

If MsgBox("Удалить документ №" + VStbd.TextMatrix(VStbd.Row, 1), vbYesNo) = vbYes Then

Ik = 0
If Dl.BOF = False Then Dl.MoveFirst
Do While Not Dl.EOF
Ik = Ik + 1
Jdite.Show
Jdite.Label1 = "-" + Str(Ik) + "-" + vbNewLine + "Подождите удаляю нач.из л/счетов> " + Str(Dl("TabNum"))
Jdite.Label1.Refresh


' Обноляю SummaB и SummaI для начислений с ненулевым сальдо
Mconn.Execute ("UPDATE Adding INNER JOIN TablDoc ON Adding.TablDoc = TablDoc.TablKod SET Adding.SummaI = 0, Adding.SummaB = 0 WHERE (((TablDoc.TablKod)=" + Str(Dl("TablKod")) + ") AND ((Adding.SaldoN)<>0))")


Mconn.Execute ("DELETE Adding.TablDoc From Adding WHERE (((Adding.TablDoc)=" + Str(Dl("TablKod")) + ")) and (Adding.SaldoN=0)")

'Расчет сальдо на начало
'MainForm.RSaldoN Dl("TabNum")
Mconn.Execute ("UPDATE Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV) SET Adding.SaldoN = [Saldo_Arh]![SK] WHERE (((Adding.KodKv)=" + VStbd.TextMatrix(VStbd.Row, 1) + "))")
'Расчет сальдо и количества

MainForm.КоличествоСальдо Str(Dl("TabNum"))

MainForm.RSaldoK Str(Dl("TabNum"))

Dl.Delete
Dl.Update
Dl.MoveNext
Loop
Mconn.Execute ("DELETE ReestrTablDoc.Cod From ReestrTablDoc WHERE (((ReestrTablDoc.Cod)=" + VStbd.TextMatrix(VStbd.Row, 1) + "))")

End If
Dl.Close
Unload Jdite
rs_kat.Requery
Set VStbd.DataSource = rs_kat
End Sub
