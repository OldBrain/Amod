VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TipKv 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���������� ����� �������"
   ClientHeight    =   6672
   ClientLeft      =   168
   ClientTop       =   468
   ClientWidth     =   5304
   FillColor       =   &H00400000&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form7"
   Picture         =   "ooo.frx":0000
   ScaleHeight     =   6672
   ScaleWidth      =   5304
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5304
      _ExtentX        =   9356
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
      Caption         =   "������� <F8>"
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
      BackColor       =   &H80000000&
      Caption         =   "��������<F4>"
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
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      _cx             =   8916
      _cy             =   10610
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
      FormatString    =   $"ooo.frx":4DF7D
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
            Picture         =   "ooo.frx":4E066
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ooo.frx":4E178
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ooo.frx":4E28A
            Key             =   "OOFL"
         EndProperty
      EndProperty
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Shortcut        =   {F4}
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   {F8}
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
      End
   End
End
Attribute VB_Name = "TipKv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection

Private Sub FG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If FG1.Col = 1 Then FG1.Editable = flexEDNone Else FG1.Editable = flexEDKbdMouse



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


Private Sub DataList1_Click()
DataList1.Refresh
End Sub

Private Sub Command1_Click()
rs_kat.UpdateBatch
Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.TipKvKod = [MainOccupant]![KV]")

Mconn.Execute ("UPDATE Tarif INNER JOIN TipKv ON Tarif.KodKV = TipKv.��� SET Tarif.NameKV = [TipKv]![Name_Kv]")



TipKv.Hide
Sprav.Show
End Sub

Private Sub Command2_Click()

Dim n, N1 As Integer
If MsgBox("�������� ����� ������?", vbYesNo) = vbYes Then
n = 0
rs_kat.MoveFirst
Do While Not rs_kat.EOF
If rs_kat("���").Value = "" Then
rs_kat.Delete
rs_kat.MoveFirst
End If
N1 = rs_kat("���").Value
If N1 > n Then n = N1
rs_kat.MoveNext
Loop

rs_kat.AddNew
rs_kat("���") = n + 1
rs_kat("NAME_KV") = "����� ��� ��������"
rs_kat.UpdateBatch
FG1.DataRefresh
rs_kat.MoveLast
End If
End Sub

Private Sub Command3_Click()
Dim DelItem As String
With rs_kat
DelItem = FG1.TextMatrix(FG1.Row, 1)
If MsgBox("�� ������ ������� �" + FG1.TextMatrix(FG1.Row, 1) + "  " + FG1.TextMatrix(FG1.Row, 2) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If rs_kat("���") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
FG1.DataRefresh
If .EOF Then .MoveLast
End If
End With

End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'If FG1.Col = 1 Then FG1.Editable = flexEDNone Else FG1.Editable = flexEDKbdMouse
rs_kat.UpdateBatch

End Sub

Private Sub FG1_Click()
'MsgBox (FG1.Cell(flexcpText))
'MsgBox (FG1.TextMatrix(FG1.Row, 1))
End Sub

Private Sub FG1_RowColChange()
FG1.Refresh
End Sub

Private Sub Form_Load()

 FG1.Editable = False

' open connection
'   Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'  mconn.Open "data/Kvartplata.mdb"
    
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
 
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
rs_kat.Open "TipKv"
Set FG1.DataSource = rs_kat


' ������������� recordset � �����
   
   FG1.FocusRect = 3
    'flexFocusSolid
    FG1.Editable = True
    FG1.DataMode = flexDMBound
    
    FG1.AutoSearch = flexSearchFromCursor
    FG1.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs_kat.Close
Mconn.Close
End Sub

Private Sub ��������_Click()
Command2_Click
End Sub

Private Sub �������_Click()
Command1_Click
End Sub

Private Sub �������_Click()
Command3_Click
End Sub
