VERSION 5.00
Begin VB.Form Sch_kat 
   Caption         =   "Form3"
   ClientHeight    =   3408
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form3"
   ScaleHeight     =   3408
   ScaleWidth      =   3744
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   1320
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   2160
      Width           =   972
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3492
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   3612
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "����� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1128
      TabIndex        =   5
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "����� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3588
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "����� ��������� �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   456
      TabIndex        =   1
      Top             =   120
      Width           =   2820
   End
End
Attribute VB_Name = "Sch_kat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Dim rsCH As ADODB.Recordset
Dim Combo_rs1 As ADODB.Recordset
Dim rsTar As ADODB.Recordset
Dim KodKat As String
Dim KodAdr As String
Dim Tarifs As String
Dim Adr1 As Integer
Dim strFindString As String

Private Sub Combo1_Change()
KodKat = Val(Combo1.Text)

End Sub

Private Sub Combo1_LostFocus()
KodKat = Val(Combo1.Text)
SCH_ET.KodKat = KodKat
��������
End Sub





Private Sub Combo2_Change()
��������
End Sub

Private Sub Combo2_GotFocus()
��������
End Sub

Private Sub Combo2_LostFocus()
Tarifs = Combo2.Text

End Sub

Private Sub Combo3_Change()
KodAdr = Combo3.Text
SCH_ET.Adr = KodAdr
End Sub

Private Sub Combo3_LostFocus()
KodAdr = Combo3.Text
SCH_ET.KodKat = KodAdr

��������
End Sub

Private Sub Command1_Click()

SCH_ET.Adr = KodAdr
SCH_ET.KodKat = KodKat
SCH_ET.Tarifs.Text = Tarifs

SCH_ET.Q.Caption = "SELECT MainOccupant.Numer AS �����, MainOccupant.kv_num AS ��, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.SummaI AS ���������, Adding.Shc_old AS [������� ����], Adding.Shc_new AS [������� �������], 0 AS [��������] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN Adding ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom WHERE (((Adding.KodKat)=" + SCH_ET.KodKat.Caption + ") AND ((Adding.Sch)='��') AND ((MainOccupant.Dom)=" + Str(Adr1) + "))ORDER BY MainOccupant.FAM"

SCH_ET.Command2.Caption = SCH_ET.Command2.Caption + " ��� ��������� ������� >" + KodKat + " �� ������ >" + KodAdr
SCH_ET.Show
Unload Me
End Sub

Private Sub Form_Load()

Set Combo_rs1 = New ADODB.Recordset
Set Combo_rs1.ActiveConnection = Mconn


' ��������� � ����� ������ � ���������� ���. ������� ���������� ��������
' ������ << rsCH("TARIFS") >> ��� ���� TarifI ���� TarfD


Set rsCH = New ADODB.Recordset
rsCH.Open ("SELECT nachisleniy.���Kategor, nachisleniy.Kategor, nachisleniy.Kod, nachisleniy.Formula, nachisleniy.Tip, nachisleniy.Sch, IIf(InStr(1,[Formula],'TarifD',0)<>0,'TarifD',IIf(InStr(1,[Formula],'TarifI',0)<>0,'TarifI','0')) AS TARIFS From Nachisleniy WHERE (((nachisleniy.Tip)='+') AND ((nachisleniy.Sch)='��'))"), Mconn
rsCH.MoveFirst
Combo1.Text = Str(rsCH("���Kategor")) + " " + rsCH("Kategor")
Combo2.Text = 0

Do While Not rsCH.EOF
'MsgBox rsCH("���Kategor")
Combo1.AddItem Str(rsCH("���Kategor")) + " " + rsCH("Kategor")
'MsgBox (rsCH("TARIFS"))
'Combo2.AddItem (rsCH("TARIFS"))
rsCH.MoveNext
Loop







' ��������� Combo3 ��� �������

Combo_rs1.Open "KLS_PODR", Mconn
'Combo3.Text = "0"
'Cl = "0"
Combo_rs1.MoveFirst
Cl = CStr(Combo_rs1("���")) & "  " & Combo_rs1("Naim_kls") & " ��� � " & Combo_rs1("Num")
Combo3.Text = Cl
Combo_rs1.MoveNext
Do While Not Combo_rs1.EOF
Cl = CStr(Combo_rs1("���")) & "  " & Combo_rs1("Naim_kls") & " ��� � " & Combo_rs1("Num")
If Trim(Cl) <> "" Then Combo3.AddItem Cl
Combo_rs1.MoveNext
Loop



End Sub

 Private Sub ��������()
 Dim Adr As ADODB.Recordset
Set Adr = New ADODB.Recordset
Set Adr.ActiveConnection = Mconn

'������� ������� ����� ������� ����� ��� ����������

Combo2.Clear
'Combo2.Text = 0
' ���������� ��� ���������� ����
Adr1 = Val(Combo3.Text)
Adr.Open ("SELECT KLS_PODR.���, KLS_PODR.Tip From KLS_PODR WHERE (((KLS_PODR.���)=" + Str(Adr1) + "))")
'MsgBox (Adr("Tip"))
'Dom_tip = Adr("Tip")

'MsgBox (Dom_tip)

 


'1. ��������� � Combo2 ����� ������ � ������� Value


Set rsTar = New ADODB.Recordset
rsTar.Open ("SELECT Tarif.KodKat, Tarif.KodDOM, Tarif.Value From Tarif WHERE (((Tarif.KodKat)=" + KodKat + ") AND ((Tarif.KodDOM)=" + Str(Dom_tip) + "))"), Mconn
If Not rsTar.EOF Then

rsTar.MoveFirst

Do While Not rsTar.EOF

' ���� ����� � �����
strFindString = Str(rsTar("Value"))
CB = SendMessage(Combo2.hwnd, CB_FINDSTRING, -1, ByVal strFindString)
If CB <> -1 Then
 '   MsgBox "Found index " + CStr(CB) + " �� ������� �� �����!!! " + strFindString
   Else
Combo2.AddItem Str(rsTar("Value"))
End If
rsTar.MoveNext



Loop
End If


'2. ��������� � Combo2 ����� ������ � ������� TarifI
Adr1 = Val(Combo3.Text)
Set rsTar = New ADODB.Recordset
rsTar.Open ("SELECT Tarif.KodKat, Tarif.KodDOM, Tarif.TarifI From Tarif WHERE (((Tarif.KodKat)=" + KodKat + ") AND ((Tarif.KodDOM)=" + Str(Dom_tip) + "))"), Mconn
If Not rsTar.EOF Then
rsTar.MoveFirst

Do While Not rsTar.EOF

' ���� ����� � �����
strFindString = Str(rsTar("TarifI"))
CB = SendMessage(Combo2.hwnd, CB_FINDSTRING, -1, ByVal strFindString)
If CB <> -1 Then
'    MsgBox "Found index " + CStr(CB) + " �� ������� �� �����!!! " + strFindString
   Else
    Combo2.AddItem Str(rsTar("TarifI"))
End If

rsTar.MoveNext
Loop






End If

'3. ��������� � Combo2 ����� ������ � ������� TarifD
Adr1 = Val(Combo3.Text)
Set rsTar = New ADODB.Recordset
rsTar.Open ("SELECT Tarif.KodKat, Tarif.KodDOM, Tarif.TarifD From Tarif WHERE (((Tarif.KodKat)=" + KodKat + ") AND ((Tarif.KodDOM)=" + Str(Dom_tip) + "))"), Mconn
If Not rsTar.EOF Then
rsTar.MoveFirst

Do While Not rsTar.EOF
' ���� ����� � �����
strFindString = Str(rsTar("TarifD"))
CB = SendMessage(Combo2.hwnd, CB_FINDSTRING, -1, ByVal strFindString)
If CB <> -1 Then
 '   MsgBox "Found index " + CStr(CB) + " �� ������� �� �����!!! " + strFindString
   Else
Combo2.AddItem Str(rsTar("TarifD"))
End If




rsTar.MoveNext
Loop
End If


End Sub



