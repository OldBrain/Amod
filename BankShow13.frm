VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BankShow13 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9096
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   11868
   ControlBox      =   0   'False
   Icon            =   "BankShow13.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   758
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "BankShow13.frx":030A
      Left            =   7920
      List            =   "BankShow13.frx":030C
      TabIndex        =   7
      Text            =   "902"
      Top             =   480
      Width           =   3852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
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
      Height          =   372
      Left            =   3240
      TabIndex        =   6
      Top             =   8640
      Width           =   1932
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11652
      _ExtentX        =   20553
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��� ������������ �/��"
            Key             =   "Key1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��� ������"
            Key             =   "Key2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����� ��������"
            Key             =   "Key3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "XL"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton Image1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5772
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   11772
      _cx             =   20764
      _cy             =   10181
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"BankShow13.frx":030E
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
      DataMode        =   2
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   255
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��� ������"
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
      Left            =   6840
      TabIndex        =   8
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   11055
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
      Left            =   0
      Picture         =   "BankShow13.frx":03F7
      Top             =   0
      Width           =   156
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
      Left            =   720
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10890
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   0
      Picture         =   "BankShow13.frx":0641
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BankShow13.frx":0D8B
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BankShow13.frx":14D5
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "BankShow13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim cRs As ADODB.Recordset
Dim RSD As ADODB.Recordset
Dim RsM As ADODB.Recordset
Dim rss As ADODB.Recordset
Dim estSc As Boolean
Dim S1 As Double
Dim Neo As String
Dim S2 As Double
Dim S3 As Double
Dim KZ As Integer

Private Sub BtnEnh11_Click()

End Sub

Private Sub Command1_Click()

' ��������� ������ � ������ ����������
rs1.MoveFirst
Set rsDocReestr = New ADODB.Recordset
rsDocReestr.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.NachCod, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc"), Mconn, adOpenKeyset, adLockPessimistic
rsDocReestr.AddNew
rsDocReestr("Coment") = Me.lblTitle + " " + Combo1.Text
Cod = rsDocReestr("Cod") ' ��� �������
rsDocReestr("Data") = rs1("���� �������")
rsDocReestr("Nach") = Combo1.Text
rsDocReestr.UpdateBatch
rsDocReestr.Close


' ��������� ������ � ������� � ���� ����������
Set RSD = New ADODB.Recordset
Set RsM = New ADODB.Recordset
Set rss = New ADODB.Recordset
rss.Open ("SELECT Settings.Neo FROM Settings"), Mconn
RSD.Open ("SELECT doc.Cod, doc.DataR, doc.KodN, doc.NameN, doc.KodKv, doc.NameKv, doc.Summa, doc.Key, doc.KeyAdding, doc.Stst, doc.Com, doc.Tip, doc.Button, doc.Dom, doc.RealData, doc.PLNOM FROM doc"), Mconn, adOpenKeyset, adLockPessimistic
RsM.Open ("SELECT MainOccupant.Numer, MainOccupant.kv_num, MainOccupant.Dom, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.OLDNUM FROM MainOccupant"), Mconn
'RsD.AddNew
Neo = rss("Neo")
rss.Close


Pod.Show
Pod.Label3.Caption = "�� ������� �����>"
Pod.Label3.Visible = False
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = KZ + 3
'MsgBox (Pod.ProgressBar1.Max)


' ���������� ������ ����� ������� ��� ����������� ������ �/��
rs1.MoveFirst
Do While Not rs1.EOF
' ��������� ��������
If rs1("� � �") <> "" Then Pod.Label1.Caption = "��� >" + rs1("� � �")
Pod.Refresh


RSD.AddNew
RSD("DataR") = rs1("���� �������")
RSD("Cod") = Cod
RSD("NameKv") = rs1("� � �")
RSD("KodN") = Val(Me.Combo1.Text)

If MainForm.ErcFile = True Then RSD("Com") = Me.lblTitle + " �� " + CStr(rs1("������ ������")) Else RSD("Com") = Me.lblTitle + " �� " + rs1("������ ������")

RSD("NameN") = Me.Combo1.Text
RSD("Summa") = rs1("�����")
RSD("Stst") = 0
RSD("Tip") = "-"
'On Error GoTo ne
'RSD("RealData") = rs1("������ ������")
'ne:
'������ ���� ������������ �������

RsM.MoveFirst
Do While Not RsM.EOF
If rs1("����") = "" Then rs1("����") = 0
If RsM("OLDNUM") = rs1("����") Then
'���� �����

If rs1("� � �") <> "" Then Pod.Label2.Caption = "�������� ������>" + rs1("� � �")
Pod.Refresh

RSD("KodKv") = RsM("Numer")
RSD("NameKv") = rs1("� � �") + "/" + RsM("FAM")
RSD("Dom") = RsM("DOM")
estSc = True
End If



RsM.MoveNext

Loop
'************************************
'���� ��� �� ������������ �����
If estSc = False Then
RSD("KodKv") = Neo
RSD("NameKv") = "�/C" + "/" + rs1("� � �") + "/" + rs1("�����")
Pod.Label3.Visible = True
Pod.Label3.Caption = Pod.Label3.Caption + rs1("� � �")
RSD("Com") = Me.lblTitle + "/" + rs1("� � �") + "/" + rs1("�����") + " �� " + rs1("������ ������")
RSD("Dom") = 1


Pod.Refresh
End If
estSc = False

Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
Pod.Refresh

rs1.MoveNext
Loop

RSD.UpdateBatch

Unload Pod

Unload Me
End Sub

Private Sub Command3_Click()
������Exel
End Sub



Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Me.Show
End Sub

Private Sub Form_Load()

Set rsDocReestr = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set cRs = New ADODB.Recordset

cRs.Open ("SELECT nachisleniy.Kod, nachisleniy.���Kategor, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy ORDER BY nachisleniy.Tip"), Mconn


' ��������� combo ��� ���� ����������
cRs.MoveFirst
'Set Combo1.DataSource = cRs
'Combo1.DataField = "Kod"

Do While Not cRs.EOF
Combo1.AddItem Str(cRs("Kod")) + "|" + cRs("Naim")
cRs.MoveNext
Loop




'��� ������

rs1.Open ("SELECT Bank.DATA AS [���� �������], Bank.LSCHET AS ����, Bank.ADR AS �����, Bank.FIO AS [� � �], Bank.SUMMA AS �����, Bank.PERIODOPL AS [������ ������] From Bank ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic
KZ = 0
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Do While Not rs1.EOF
If MainForm.ErcFile = False Then S1 = S1 + rs1("�����")
rs1.MoveNext
KZ = KZ + 1
Loop
End If


' "����� ��������"
rs2.Open ("SELECT Bank.DATA, Bank.LSCHET AS ����, Bank.ADR AS �����, Bank.FIO AS [� � �], Bank.SUMMA AS �����, Bank.PERIODOPL AS [������ ������] From Bank WHERE (((Bank.LSCHET) Is Not Null Or (Bank.LSCHET)='0')) ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic

If rs2.RecordCount > 0 Then
rs2.MoveFirst
Do While Not rs2.EOF
S2 = S2 + rs2("�����")
rs2.MoveNext
Loop
End If

' ��� ������������
rs3.Open ("SELECT Bank.DATA AS [���� ������], Bank.SUMMA AS �����, Bank.FIO AS [� � �], Bank.ADR AS �����, Bank.LSCHET AS [�/����], Bank.PERIODOPL AS [������ ������] FROM Bank LEFT JOIN MainOccupant ON Bank.LSCHET = MainOccupant.OLDNUM WHERE (((MainOccupant.OLDNUM) Is Null))"), Mconn, adOpenKeyset, adLockPessimistic

If rs3.RecordCount > 0 Then
rs3.MoveFirst
Do While Not rs3.EOF
If MainForm.ErcFile = False Then S3 = S3 + rs3("�����")
rs3.MoveNext
Loop
End If

MakeWindow Me, True
FG1.Width = Me.Width / 15.40107
FG1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
Command1.Top = Image1.Top

'**************************************************************
'****************************************************************

'lblTitle = "������ ������ �� �����. ���� > " + BankImport.File1.FileName
'Label1.Caption = "�������� ����� >" + BankImport.File1.FileName + ". ��� ����������� ������� <<�����>>"

'Label1.Caption = TabStrip1.SelectedItem + " �� ����� >" + Str(S3)

lblTitle.Caption = lblTitle.Caption + "����� ������������  > " + Str(BankShow.SummI)
'TabStrip1.Index = 1


Set FG1.DataSource = rs3
End Sub


Private Sub Image1_Click()

'��������� Realdata



'Msg.Show vbModal
'Msg.Label1.Caption = "������ ��������� � ������ ���������� ������ �" + Str(BankShow.Cod) + vbNewLine + "����� ������� �����=" + Str(BankShow.SummI) + vbNewLine + "����������� �����=" + Str(BankShow.SummI - S3) + vbNewLine + "����������=" + Str(Round(S3, 2))
'Msg.Label1.Refresh

 'Unload ReestrDoc

'ReestrDoc.Fg.Refresh


Unload Me
Msg.Show
Msg.Label1.Caption = "������ ��������� � ������ ���������� ������ �" + Str(BankShow.Cod) + vbNewLine + "����� ������� �����=" + Str(BankShow.SummI) + vbNewLine + "����������� �����=" + Str(Round(BankShow.SummI - S3, 2)) + vbNewLine + "����������=" + Str(Round(S3, 2))
Msg.Label1.Refresh


Unload BankShow

ReestrDoc.Show
ReestrDoc.Enabled = True
End Sub
Sub ������Exel()
   Const ��������� = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim i As Long, j As Long, K As Long, r������ As String
   Dim v As Variant
   
   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
   
   r������ = "A" & (��������� + 1) & ":" & XCol_(FG1.Cols - 1) & FG1.Rows + ���������
   ReDim v(FG1.Rows, FG1.Cols) '����� �������
   
   If FG1.Rows > 0 Then
            For co = 1 To FG1.Cols - 1
         For rw = 0 To FG1.Rows - 1

             v(rw, co) = FG1.TextMatrix(rw, co)
             
         Next rw
         Next co
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


Private Sub imgTitleHelp_Click()
Form2.Label1 = "   ������ ������ ������������ �����, �� ����� ������ ��������������� ������. ��� ���� ������������� ��� ���������������� ��������� ������������� ������." + vbNewLine + "   ����� ����, �� ������ ��������� ������ � XL ��� ���������� ������"
Form2.Show
End Sub

Private Sub imgTitleMain_Click()
ChangeState Me
End Sub

Private Sub lblTitle_Click()
ChangeState Me
End Sub
Private Sub Form_Resize()
FG1.Width = Me.Width / 15.40107
   FG1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
'BtnEnh1.Top = Image1.Top

Command3.Left = Image1.Left + Image1.Width
End Sub


Private Sub TabStrip1_Click()



If TabStrip1.SelectedItem = "��� ������" Then
Set FG1.DataSource = rs1
Me.Command1.Enabled = False
Label1.Caption = TabStrip1.SelectedItem + " �� ����� >" + Str(S1)
End If

If TabStrip1.SelectedItem = "����� ��������" Then
Me.Command1.Enabled = True
Set FG1.DataSource = rs2
Label1.Caption = TabStrip1.SelectedItem + " �� ����� >" + Str(S2)
End If

If TabStrip1.SelectedItem = "��� ������������ �/��" Then
Me.Command1.Enabled = False
Set FG1.DataSource = rs3
Label1.Caption = TabStrip1.SelectedItem + " �� ����� >" + Str(S3)
End If

'
'SelectedItem

End Sub

