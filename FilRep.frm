VERSION 5.00
Begin VB.Form FilRep 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3804
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   5676
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3804
   ScaleWidth      =   5676
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "�� ���������"
      Height          =   435
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "����������� ������"
      Height          =   435
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��������� ������"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Text            =   "500"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "FilRep.frx":0000
      Left            =   720
      List            =   "FilRep.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "���� �����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "FilRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdrConn As ADODB.Recordset
'Dim mconn As ADODB.Connection




Private Sub Combo1_Change()

'If Not Adrconn("Num") Then Combo2.Text = Adrconn.Fields("NAIM_KLS") + " ��� � " + Adrconn("Num")


End Sub

Private Sub Command1_Click()


Dim sq As String
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 2))
sq = ""


'MsgBox (fil)
'MsgBox (Combo1.Text)

'sq = "SELECT [KLS_PODR]![NAIM_KLS]+" + Chr(34) + " ��� �" + Chr(34) + "+Str([KLS_PODR]![Num]) AS �����, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, [Adding]![SaldoK]/[Adding]![Kol] AS ����, MainOccupant.Dom FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� Where ((([Adding]![SaldoK] / [Adding]![Kol]) > " + Text1 + ") And ((MainOccupant.Dom) = " + Str(Val(Combo1.Text)) + ")) ORDER BY MainOccupant.kv_num"
'sq = "SELECT [KLS_PODR]![NAIM_KLS]+" + Chr(34) + " ��� �" + Chr(34) + "+Str([KLS_PODR]![Num]) AS �����, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, round(([Adding]![SaldoK]/[Adding]![Kol]),2) AS ���� FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� Where ((([Adding]![SaldoK] / [Adding]![Kol]) > " + Text1 + ") And ((MainOccupant.Dom) = " + Str(Val(Combo1.Text)) + ")) ORDER BY MainOccupant.kv_num"

If Option1 = True Then

sq = "SELECT MainOccupant.Numer, MainOccupant.kv_num as [�� �] , MainOccupant.FAM as �������, MainOccupant.IM as ���, MainOccupant.OT as ��������, Sum([Adding]![SaldoK]/[Adding]![Kol]) AS ����, KLS_PODR.��� FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY MainOccupant.Numer, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.��� Having (((Sum([Adding]![SaldoK] / [Adding]![Kol])) > " + Text1 + ") And ((KLS_PODR.���) = " + Str(fil) + ")) ORDER BY MainOccupant.kv_num"




End If

If Option2 = True Then
MsgBox Text1 + "  " + Str(fil)

sq = "SELECT " + Chr(34) + "�� � " + Chr(34) + "+ [MainOccupant]![kv_num] +" + Chr(34) + " " + Chr(34) + "+ [MainOccupant]![FAM] +" + Chr(34) + " " + Chr(34) + "+ [MainOccupant]![IM]+" + Chr(34) + " " + Chr(34) + "+MainOccupant.OT AS [� �� ������� ��� ��������], Adding.NameKat AS [��������� �������], Sum([Adding]![SaldoK]/[Adding]![Kol]) AS ����, KLS_PODR.��� FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.NameKat, KLS_PODR.��� Having (((Sum([Adding]![SaldoK] / [Adding]![Kol])) > " + Text1 + ") And ((KLS_PODR.���) = " + Str(fil) + ")) ORDER BY MainOccupant.kv_num"
End If


If Option3 = True Then

'Analizlgot.G = 4
'sq = "SELECT MainOccupant.podyezd as �������, MainOccupant.Numer, MainOccupant.kv_num AS [�� �], MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Sum([Adding]![SaldoK]/[Adding]![Kol]) AS ����, KLS_PODR.��� FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY MainOccupant.podyezd, MainOccupant.Numer, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.��� Having (((Sum([Adding]![SaldoK] / [Adding]![Kol])) > " + Text1 + ") And ((KLS_PODR.���) = " + Str(fil) + ")) ORDER BY MainOccupant.kv_num"
'
sq = "SELECT MainOccupant.podyezd,  MainOccupant.kv_num AS [�� �], MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Sum([Adding]![SaldoK]/[Adding]![Kol]) AS ���� FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY MainOccupant.podyezd, MainOccupant.Numer, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.��� Having (((Sum([Adding]![SaldoK] / [Adding]![Kol])) > " + Text1 + ") And ((KLS_PODR.���) = " + Str(fil) + ")) ORDER BY MainOccupant.kv_num"
'Analizlgot.�� 1
End If

If Combo1.Text = "������ �����" Then
MsgBox ("������ �����")
Exit Sub
End If

'������� ������ �� ��������� ��� ������ �����������

Reports.sq = sq
Analizlgot.Fg1.Cols = 4




Analizlgot.Titl = MainForm.Label3 + vbNewLine + " ��������� ��������� ����������� �� ������:" + Combo1.Text + " �� " + MonthName(Month(MainForm.DR)) + " " + Str(Year(MainForm.DR)) + " �., ������� ���� �� ����� ������� ����� " + Text1 + " ���."



Analizlgot.Show
Analizlgot.Fg1.Subtotal flexSTClear
Analizlgot.Fg1.DataRefresh

Analizlgot.Dol.Visible = True


If Option1 = True Then
Analizlgot.�� 1

' ���������� ���������� �������
Analizlgot.Fg1.Cols = 7

' ������ ������� � �������
Analizlgot.Fg1.ColHidden(1) = True




End If
If Option2 = True Then
Analizlgot.�� 2
Analizlgot.Fg1.Cols = 4

End If

If Option3 = True Then
'Analizlgot.Fg1.Cols = 8
Analizlgot.�� 2


End If

Unload FilRep
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
MainMenu.Enabled = True
MainMenu.Show
End Sub

Private Sub Form_Load()
Option1 = True

'Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"


Set AdrConn = New ADODB.Recordset
Set AdrConn.ActiveConnection = Mconn
AdrConn.CursorType = adOpenStatic
AdrConn.LockType = adLockBatchOptimistic


'Adrconn.Open ("KLS_PODR")
AdrConn.Open ("SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.�������������, KLS_PODR.���� From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Combo1.Text = "������ �����"


AdrConn.MoveFirst
Combo1.AddItem "��� ����"
Do While Not AdrConn.EOF
If AdrConn("���") <> -1 Then
Combo1.AddItem Str(AdrConn("���")) + " " + AdrConn("NAIM_KLS") + " ��� � " + AdrConn("Num")
End If
AdrConn.MoveNext
Loop
End Sub
