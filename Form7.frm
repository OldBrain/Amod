VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����������� ������ ��� ����?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim PeresLgot As ADODB.Recordset
'Dim mconn As ADODB.Connection
'Set mconn = New ADODB.Connection

'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"

Set PeresLgot = New ADODB.Recordset

PeresLgot.Open ("SELECT Adding.Key, Adding.LgotaP, Adding.ObPl, Adding.Lig From Adding WHERE (((Adding.Lig)=" + Chr(34) + "��" + Chr(34) + "))"), mconn, adOpenForwardOnly, adLockPessimistic


'�1  �������� ��� Update
'adOpenKeyset , adLockOptimistic

'�2  �������� ��� Update ����������� ������� ��� �1
'adOpenKeyset , adLockPessimistic

'� 3 ����������  ������������ ������ ����� UpdateBath ������ ������
'adOpenKeyset, adLockBatchOptimistic

'� 4 adOpenForwardOnly, adLockPessimistic
' ���������� ������ �������� ������� ��� � 2




PeresLgot.MoveFirst
N = 1
Do While Not PeresLgot.EOF







MainForm.Ostatok = PeresLgot.Fields("Obpl").Value
MainForm.II = 0
MainForm.Pi = 0
MainForm.�� PeresLgot.Fields("Key").Value, True



If MainForm.������� = True Then

MainForm.Ostatok = PeresLgot.Fields("Obpl").Value
MainForm.Pi = 0
MainForm.II = 0
MainForm.�� PeresLgot.Fields("Key").Value, False

End If


MainForm.ViborLLg PeresLgot.Fields("Key").Value
PeresLgot.Fields("LgotaP").Value = MainForm.PrZ
'Str (N)
'MainForm.PrZ
'PeresLgot.Update

Label1 = Str(N)
Label1.Refresh

N = N + 1


'If N > N - (N / 100) = 0 Then Form7.Refresh
PeresLgot.MoveNext
Loop

PeresLgot.UpdateBatch

Label1 = "�������� ����� �������"
Form7.Refresh
End Sub

