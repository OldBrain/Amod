VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Arhiv_all 
   Caption         =   "�������� ������ �����"
   ClientHeight    =   7776
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12744
   LinkTopic       =   "Form3"
   ScaleHeight     =   7776
   ScaleWidth      =   12744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "���������� ����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1800
      TabIndex        =   3
      Top             =   7320
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "XL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   960
      TabIndex        =   2
      Top             =   7320
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   852
   End
   Begin VSFlex8Ctl.VSFlexGrid fg1 
      Height          =   6852
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   12612
      _cx             =   22246
      _cy             =   12086
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
      BackColorSel    =   -2147483645
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483647
      GridColorFixed  =   4194432
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   43
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Arhiv_all.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   1
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
      ShowComboButton =   0
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
      AccessibleRole  =   50
   End
End
Attribute VB_Name = "Arhiv_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARS As ADODB.Recordset


Private Sub Command1_Click()
'������� ��������� ������
zRep = "������ �/������ � ������� �������� ������ �������� ������ �� ������������� ��������� ������ ����������� ������ �������"


PrintW.Show
     With PrintW.VP
     
        PrintW.VP.StartDoc
        .FontSize = 12
        .Paragraph = zRep + vbNewLine + "_________________________________________________________________"
        .Paragraph = ""
        
        .FontSize = 8
        .RenderControl = fg1.hwnd
        .EndDoc
        
       End With


End Sub

Private Sub Command2_Click()
'MsgBox ("�������� �������� �������������")
'Exit Sub

Pod.Show
Pod.Label1 = "��������� ���� ������� ������ � XL"

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

fg1.Subtotal flexSTClear
For I = 250 To 500
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

fg1.DataRefresh
For I = 500 To 750
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

������Exel
For I = 750 To 1000
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

Unload Pod


End Sub

Private Sub Command3_Click()
'Set RS = New Recordset

If MsgBox("��������������� ����� ������ �� ������ �������? ���  ����� ������������ ������ ����� �������", vbYesNo) = vbYes Then


Pod.Show 0
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = Me.fg1.Rows + 10

For R = 1 To fg1.Rows - 1
n = fg1.TextMatrix(R, 1)
 'MsgBox (n)
s = fg1.TextMatrix(R, 8)
s = Replace(s, ",", ".")
K = fg1.TextMatrix(R, 6)
Mconn.Execute ("UPDATE Adding SET Adding.SaldoN = " + s + " WHERE (((Adding.KodKv)=" + n + ") AND ((Adding.KodKat)=" + K + "))")







Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
Pod.Refresh
Next


'��������� ������������� ������

'Mconn.Execute ("INSERT INTO Adding ( KodKv, KodKat, SaldoN, KodN, TablDoc )SELECT GroupArh.KodKv, GroupArh.KodKat, GroupArh.SaldoK, GroupNachislen.[Min-Kod], GroupAdding.KodKv FROM GroupNachislen INNER JOIN (GroupArh LEFT JOIN GroupAdding ON (GroupArh.KodKat = GroupAdding.KodKat) AND (GroupArh.KodKv = GroupAdding.KodKv)) ON GroupNachislen.���Kategor = GroupArh.KodKat WHERE (((GroupAdding.KodKv) Is Null))")


Mconn.Execute ("add_adding")


'n = Val(Filter.FG.TextMatrix(R, 0))




'������ ���� Saldo_Arh
Mconn.Execute ("DELETE Saldo_Arh.* FROM Saldo_Arh")
'��������� ������ � Saldo_arh ��� ����������� ��������
Mconn.Execute ("INSERT INTO Saldo_arh ( KodKV, KodKat, SK ) SELECT Arh_Rep_All.KodKv, Arh_Rep_All.KodKat, Sum(([Arh_Rep_All]![SaldoK]*1000)/[Arh_Rep_All]![Kol])/1000 AS Sk From Arh_Rep_All GROUP BY Arh_Rep_All.KodKv, Arh_Rep_All.KodKat")



'RS.Close
MsgBox ("������ ������� �����������. �� ������� ����������� ����������� ��� ������� �����!")
Unload Pod

Unload Me
Else

End If
End Sub

Private Sub Form_Load()
Dim D, D0 As Date
Dim RS As ADODB.Recordset

'���� ���� �������
Set RS = New Recordset
RS.Open "SELECT Settings.* FROM Settings", Mconn
D0 = RS("TekData")
RS.Close
'������� ���������� �����
D = DateAdd("m", -1, D0)
'������ ��� ����� ������ bakName
    M = MonthName(Month(D), True)
    G = Trim(Str(Year(D)))
    
    bakName = App.Path & "\data\Arhiv\" & G + M + ".amd"
'MsgBox (bakName)

'������ ����
Mconn.Execute ("DELETE Arh_Rep_All.* FROM Arh_Rep_All")

' �������� ������ �� ������ ����������� ������ � ������� ����
Mconn.Execute ("INSERT INTO arh_rep_all SELECT Adding.* FROM Adding IN '" + bakName + "'")

'���������� ������ � Adding � ����� � arh_rep_all
'RS.Open ("SELECT Kategor.Name_Kategor AS ���������, KLS_PODR.NAIM_KLS AS �����, Saldo_all.KodKv AS �����1, MainOccupant.OLDNUM AS �����2, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.kv_num AS ��, Saldo_all.������ AS [������ �� ������ �������], Saldo_arh_all.������ AS [������ �� ������ ������� ������], [Saldo_arh_all.������]-[Saldo_all.������] AS ����������� FROM (((Saldo_all INNER JOIN Saldo_arh_all ON (Saldo_all.KodKat = Saldo_arh_all.KodKat) AND (Saldo_all.KodKv = Saldo_arh_all.KodKv)) INNER JOIN MainOccupant ON Saldo_arh_all.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) INNER JOIN Kategor ON Saldo_arh_all.KodKat = Kategor.��� Where ((([Saldo_arh_all.������] - [Saldo_all.������]) <> 0)) ORDER BY [Saldo_arh_all.������]-[Saldo_all.������]"), Mconn

'RS.Open ("SELECT Kategor.Name_Kategor AS ���������, KLS_PODR.NAIM_KLS AS �����, Saldo_all.KodKv AS �����1, MainOccupant.OLDNUM AS �����2, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.kv_num AS ��, Saldo_all.������ AS [������ �� ������ �������], Saldo_arh_all.������ AS [������ �� ����� ������� ������], [Saldo_arh_all.������]-[Saldo_all.������] AS �����������, Saldo_arh_all.KodKat FROM (((Saldo_all INNER JOIN Saldo_arh_all ON (Saldo_all.KodKv = Saldo_arh_all.KodKv) AND (Saldo_all.KodKat = Saldo_arh_all.KodKat)) INNER JOIN MainOccupant ON Saldo_arh_all.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) INNER JOIN Kategor ON Saldo_arh_all.KodKat = Kategor.��� Where ((([Saldo_arh_all.������] - [Saldo_all.������]) <> 0)) ORDER BY [Saldo_arh_all.������]-[Saldo_all.������]"), Mconn
RS.Open ("SELECT KLS_PODR.NAIM_KLS, Adding.KodKv, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.KodKat, Kategor.Name_Kategor, Arh_Rep_All.SaldoK AS [������� �����], Adding.SaldoN AS [������� �����], [Arh_Rep_All]![SaldoK]-[Adding]![SaldoN] AS ����������� FROM (((Arh_Rep_All INNER JOIN Adding ON (Arh_Rep_All.KodKat = Adding.KodKat) AND (Arh_Rep_All.KodKv = Adding.KodKv)) INNER JOIN MainOccupant ON Arh_Rep_All.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) INNER JOIN Kategor ON Adding.KodKat = Kategor.��� GROUP BY KLS_PODR.NAIM_KLS, Adding.KodKv, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.KodKat, Kategor.Name_Kategor, Arh_Rep_All.SaldoK, Adding.SaldoN, [Arh_Rep_All]![SaldoK]-[Adding]![SaldoN] Having ((([Arh_Rep_All]![SaldoK] - [Adding]![SaldoN]) <> 0)) ORDER BY [Arh_Rep_All]![SaldoK]-[Adding]![SaldoN] DESC; ")

'****************************
 
fg1.AllowUserResizing = flexResizeBoth
fg1.Sort = flexSortGenericAscending
'fg1.Cols = G
fg1.ExplorerBar = flexExMove
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.MergeCol(fg1.Cols - 1) = False
fg1.MergeCol(-1) = True
'�����������
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.Refresh
fg1.Sort = flexSortGenericAscending
fg1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
fg1.RowHeight(0) = 500
fg1.WordWrap = True
fg1.Cell(flexcpAlignment, 0, 0, 0, fg1.Cols - 1) = flexAlignCenterCenter


Set fg1.DataSource = RS

End Sub
Sub ������Exel()
   Const ��������� = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim I As Long, J As Long, K As Long, r������ As String
   Dim v As Variant
   
   
'rs.CursorType = adOpenStatic
'rs.LockType = adLockReadOnly



   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
'   Set rs = Rs_kat.Clone

 'Set rs = Rs_kat
  ' rs.Filter = Rs_kat.Filter
   'rs.Sort = Rs_kat.Sort
 ' k = FG1.Rows - 1
  ' Rs_kat.MoveLast
'   r������ = "A" & (��������� + 1) & ":" & XCol_(k) & Rs_kat.RecordCount + ���������
   
   r������ = "A" & (��������� + 1) & ":" & XCol_(fg1.Cols - 1) & fg1.Rows + ���������
   ReDim v(fg1.Rows, fg1.Cols) '����� �������
'   If rs.RecordCount > 0 Then

   'If Rs_kat.RecordCount > 0 Then
   If fg1.Rows > 0 Then
    '  Rs_kat.MoveFirst
      'i = 0
      'Do Until Rs_kat.EOF
         For co = 1 To fg1.Cols - 1
         For rw = 0 To fg1.Rows - 1
             'v(i, j) = Rs_kat.Fields(j).Value
             v(rw, co) = fg1.TextMatrix(rw, co)
             
         Next rw
         Next co
         'Rs_kat.MoveNext
      'Loop
      ex1.Visible = True   '��� �����
      
      ws.Range(r������) = v
      
 End If
End Sub

Function XCol_(ByVal Column_ As Long) As String
    If (Column_ < 0) Then Column_ = 0
    If (Column_ < 26) Then
        XCol_ = Chr(Column_ + Asc("A"))
    ElseIf (Column_ < 676) Then
        XCol_ = Chr((Column_ \ 26) + Asc("A") - 1) & Chr((Column_ Mod 26) + Asc("A"))
    Else
        XCol_ = "ZZ"
    End If
End Function

