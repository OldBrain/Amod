VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BestConsulting Soft 2004  ���: 38-17-55   Ver. 1.00W"
   ClientHeight    =   5700
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   6360
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "2. ��� �2 (�������� ������)"
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   4920
      Width           =   5175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "������ ����������"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1. ��� �1 (����� ������ �� ���� Infin)"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�����"
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
      Left            =   720
      TabIndex        =   10
      Top             =   5280
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���� � �����"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   4200
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2160
      Left            =   120
      Picture         =   "Form_kv.frx":0000
      ScaleHeight     =   108
      ScaleMode       =   2  'Point
      ScaleWidth      =   156
      TabIndex        =   5
      Top             =   960
      Width           =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Text            =   "2004"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "����� �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "��� �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "������ ������ ��  ""����������"" �������� �����"
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "���� � �����:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������� ���������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ����� 
         Caption         =   "����� ������������ ����"
         Index           =   45
      End
   End
   Begin VB.Menu Mes 
      Caption         =   "����� �������"
      Index           =   0
      NegotiatePosition=   2  'Middle
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   1
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Index           =   2
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Index           =   3
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   4
      End
      Begin VB.Menu ��� 
         Caption         =   "���"
         Index           =   5
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Index           =   6
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Index           =   7
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   8
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Index           =   9
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Index           =   10
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   10
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Index           =   12
      End
      Begin VB.Menu ����� 
         Caption         =   "��������� ������"
         Index           =   99
      End
   End
   Begin VB.Menu ���������� 
      Caption         =   "���������"
      Index           =   101
      Begin VB.Menu ����������_������ 
         Caption         =   "���������� ������ (F12)"
         Index           =   104
         Shortcut        =   {F12}
      End
      Begin VB.Menu �����_�������� 
         Caption         =   "��������� ��������"
         Index           =   103
      End
      Begin VB.Menu ���� 
         Caption         =   "���� � ��������� ����������"
         Index           =   102
      End
   End
   Begin VB.Menu ���������� 
      Caption         =   "� ���������"
      Index           =   999
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private P, God, M, Pt, PtSet, Pt1 As String
Dim Er As Label











Private Sub Command1_Click()
Call MakePt(Pt)
Form1.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
Form1.Hide
'Dialog.Drive1.Drive = "c:\"
'Dialog.Dir1.Path = Dialog.Drive1.Drive
'Dialog.Dir1.Refresh
Dialog.Show
End Sub

Private Sub Command3_Click()
Form1.Hide
Dialog.Hide
'End

End Sub

Private Sub Command4_Click()

Dim Cn As ADODB.Connection
Set Cn = New ADODB.Connection

'**********���� � ������******
Call MakePt(Pt)
'******************




Mass.Show
Mass.Command1.Visible = False
Mass.Command2.Visible = False

Mass.Refresh
Mass.Label2.Caption = "���������, ���� ��������� ������"
On Error GoTo Er

'****************** ��������� ���� � ������ � Path
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog= DATA\"
Pt_Set = Dialog.Dir1.Path

'Pt_Set1 = DataList1.DataSource
'Pt_Set1 = DataEnvironment1.FILE_PAT

Cn.Execute "UPDATE Pt SET PAT = '" & Pt_Set & "'"
Cn.Close

'************** L*.DBF �������� ���������� ����������, � ��� ������� �������

'������ ����������� ����� Kls_priv.dbf

Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog= DATA\"
Cn.Execute "Delete DATA\Kls_priv.* from DATA\Kls_priv.dbf"
a = Dialog.Dir1.Path
a = a + "\G" + Form1.Text1 + "\DBF\KLS_PRIV.dbf"

'MsgBox A
'DataEnvironment1.VO
Cn.Execute "INSERT INTO DATA\Kls_priv.dbf ( KLS_PRIV.N_KLS, KLS_PRIV.NAIM_KLS, KLS_PRIV.LPROCSPACE, KLS_PRIV.LPROCOTP, KLS_PRIV.LPROCCOM, KLS_PRIV.LPROCTECH, KLS_PRIV.LPROCMUSOR, KLS_PRIV.USESPACE, KLS_PRIV.USEOTP, KLS_PRIV.USECOM, KLS_PRIV.USETECH, KLS_PRIV.USEMUSOR ) SELECT * From " & a
'cn.Execute "INSERT INTO DataEnvironment1.VO ( KLS_PRIV.N_KLS, KLS_PRIV.NAIM_KLS, KLS_PRIV.LPROCSPACE, KLS_PRIV.LPROCOTP, KLS_PRIV.LPROCCOM, KLS_PRIV.LPROCTECH, KLS_PRIV.LPROCMUSOR, KLS_PRIV.USESPACE, KLS_PRIV.USEOTP, KLS_PRIV.USECOM, KLS_PRIV.USETECH, KLS_PRIV.USEMUSOR ) SELECT * From " & A

Cn.Close


'������ ����������� ����� (������) Kls_podr.dbf

Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog= DATA\"
Cn.Execute "Delete DATA\Kls_PODR.* from DATA\Kls_PODR.dbf"
a = Dialog.Dir1.Path
a = a + "\G" + Form1.Text1 + "\DBF\KLS_PODR.dbf"
'MsgBox A
Cn.Execute "INSERT INTO DATA\Kls_podr.dbf ( KLS_PODR.N_KLS, KLS_PODR.NAIM_KLS, KLS_PODR.KOD_KLS ) SELECT * From " & a
Cn.Close

'������ ����������� ����� ������� Kls_hab.dbf

Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog= DATA\"
Cn.Execute "Delete DATA\Kls_HAB.* from DATA\Kls_hab.dbf"
a = Dialog.Dir1.Path
a = a + "\G" + Form1.Text1 + "\DBF\KLS_HAB.dbf"
Cn.Execute "INSERT INTO DATA\Kls_hab.dbf (  KLS_HAB.N_KLS, KLS_HAB.NAIM_KLS ) SELECT * From " & a
Cn.Close

'������ ����������� ���������� Kls_vo.dbf

Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog= DATA\"
Cn.Execute "Delete DATA\Kls_vo.* from DATA\Kls_vo.dbf"
a = Pt + "Kls_vo.dbf"
Cn.Execute "INSERT INTO DATA\Kls_vo.dbf (  KLS_HAB.N_KLS, KLS_HAB.NAIM_KLS ) SELECT * From " & a
Cn.Close


'GoTo sss

'**************������� ������ �� ����� L*.DBF  ***********************
   While Len(s) > 0
   s = Null
   Wend
   
   n = 99
   K = 0
   s = Dir(Pt + "L*.dbf")
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=DATA\"
Cn.Execute "Delete All_L.* from All_L.dbf"
Cn.Execute "Update data\pt.dbf Set PAT = '" & Pt & "'"
Cn.Close

While Len(s) > 0
Mass.Label2.Caption = "���� >" + s
Mass.Label3.Caption = Str(K)
Mass.Refresh
If (s <> "L-1.DBF" And s <> "KLS_FORM.DBF" And s <> "KLS_VO.DBF") Then
n = n + 1
K = K + 1
Sp = ""
'****************************
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=" & Pt

'cn.Execute "Insert Into data\All_L.dbf (TABN, NUMHABIT, FAM, IM, OT, NLODGER, NLODGERF, NROOM, COMSPACE, HABSPACE, PRIVILEGE, HABITATE, BIRTHDAY, NORDER, DATAORDER, ORGORDER, COSTBUILD, COSTFULL, KITCHSPACE, BATHSPACE, CORRSPACE, TOILSPACE, BALCSPACE, NFAMILY, DATARECEIV, PASSPORT, TELEPHONE, LDOK, LDATEBEG, LDATEEND, NAPARTMENT, DATEIN, DATEOUT, SROKIN, FLOOR, SUBFAM, SUBCOM, SUBDOK, COMM) Select TABN, NUMHABIT, FAM, IM, OT, NLODGER, NLODGERF, NROOM, COMSPACE, HABSPACE, PRIVILEGE, HABITATE, BIRTHDAY, NORDER, DATAORDER, ORGORDER, COSTBUILD, COSTFULL, KITCHSPACE, BATHSPACE, CORRSPACE, TOILSPACE, BALCSPACE, NFAMILY, DATARECEIV, PASSPORT, TELEPHONE, LDOK, LDATEBEG, LDATEEND, NAPARTMENT, DATEIN, DATEOUT, SROKIN, FLOOR, SUBFAM, SUBCOM, SUBDOK, COMM from " & s & " Where trim(TABN) <> null"
'A = "Insert Into data\All_L.dbf ([TABN], [VID]) Select [TABN], [VID] from " & s & " Where trim(TABN) <> null"
'MsgBox (A)
'cn.Execute "UPDATE " & s & " SET Money = 0"
Cn.Execute "Insert Into data\All_L.dbf ([TABN], [VID], [Money]) Select [TABN], [VID], 0 from " & s & " Where trim(TABN) <> null"

'cn.Execute "UPDATE " & s & " SET Money = 0"
Cn.Execute "UPDATE data\all_l SET all_l.DOM = '" & Trim(Str(n)) & "' WHERE(all_l.Money = 0)"
Cn.Execute "UPDATE data\all_l SET all_l.Money = 1"
Cn.Close
'*****************************
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=data\"
Cn.Execute "UPDATE data\all_l SET all_l.DOM = '" & Right(s, Len(s) - 1) & "' WHERE(all_l.DOM=('" & Trim(Str(n)) & "'))"
Cn.Close
End If
s = Dir()
Wend
'/////////////////////////////////////////////////////



'sss:





'**************������� ������ �� ����� Z*.DBF  ***********************
   'Dim s As String
   n = 99
   K = 0
   s = Dir(Pt + "Z*.dbf")
   
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog= DATA\"
Cn.Execute "Delete DATA\All_Z.* from DATA\All_Z.dbf"
'cn.Execute "Update data\pt.dbf Set PAT = '" & Pt & "'"
Cn.Close

On Error GoTo Er
While Len(s) > 0
Mass.Label2.Caption = "���� >" + s
Mass.Label3.Caption = Str(K)
Mass.Refresh
If s <> "Z-1.DBF" Then
n = n + 1
K = K + 1
Sp = ""
'****************************
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=" & Pt
'A = "INSERT INTO DATA\ALL_Z.DBF ( TABN, VID, [MONTH], [DAY], [MONEY], ZAK, CHET, DATEPAY, VALMETER, PRIM, PENI ) SELECT TABN, VID, [MONTH], [DAY], [MONEY], ZAK, CHET, DATEPAY, VALMETER, PRIM, PENI FROM " & s & " Where trim(TABN) <> null"
'MsgBox (A)
Cn.Execute "INSERT INTO DATA\ALL_Z.DBF ( TABN, VID, [MONTH], [DAY], [MONEY], ZAK, CHET, DATEPAY, VALMETER, PRIM, PENI ) SELECT TABN, VID, [MONTH], [DAY], [MONEY], ZAK, CHET, DATEPAY, VALMETER, PRIM, PENI FROM " & s & " Where trim(TABN) <> null"
Cn.Execute "UPDATE " & s & " SET PENI = 0"
Cn.Execute "UPDATE data\all_Z SET all_Z.DOM = '" & Trim(Str(n)) & "' WHERE(all_Z.PENI < " & n & ")"
Cn.Execute "UPDATE data\all_Z SET all_Z.PENI = " & (n + 10)
Cn.Close
'*****************************
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=data\"

'MsgBox (Right(s, Len(s) - 1))
'MsgBox (Str(N))

Cn.Execute "UPDATE data\all_Z SET all_Z.DOM = '" & Right(s, Len(s) - 1) & "' WHERE(all_Z.DOM=('" & Trim(Str(n)) & "'))"
Cn.Close
End If
s = Dir()
Wend
sss:
'**************������� ������ �� ����� �*.DBF  ***********************
   While Len(s) > 0
   s = Null
   Wend
   
   n = 99
   K = 0
   s = Dir(Pt + "K*.dbf")
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=DATA\"
Cn.Execute "Delete All_K.* from All_k.dbf"
Cn.Execute "Update data\pt.dbf Set PAT = '" & Pt & "'"
Cn.Close

While Len(s) > 0
Mass.Label2.Caption = "���� >" + s
Mass.Label3.Caption = Str(K)
Mass.Refresh
If (s <> "K-1.DBF" And s <> "KLS_FORM.DBF" And s <> "KLS_VO.DBF") Then
n = n + 1
K = K + 1
Sp = ""
'****************************
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=" & Pt

'cn.Execute "Insert Into data\All_k.dbf (TABN, NUMHABIT, FAM, IM, OT, NLODGER, NLODGERF, NROOM, COMSPACE, HABSPACE, PRIVILEGE, HABITATE, BIRTHDAY, NORDER, DATAORDER, ORGORDER, COSTBUILD, COSTFULL, KITCHSPACE, BATHSPACE, CORRSPACE, TOILSPACE, BALCSPACE, NFAMILY, DATARECEIV, PASSPORT, TELEPHONE, LDOK, LDATEBEG, LDATEEND, NAPARTMENT, DATEIN, DATEOUT, SROKIN, FLOOR, SUBFAM, SUBCOM, SUBDOK, COMM) Select TABN, NUMHABIT, FAM, IM, OT, NLODGER, NLODGERF, NROOM, COMSPACE, HABSPACE, PRIVILEGE, HABITATE, BIRTHDAY, NORDER, DATAORDER, ORGORDER, COSTBUILD, COSTFULL, KITCHSPACE, BATHSPACE, CORRSPACE, TOILSPACE, BALCSPACE, NFAMILY, DATARECEIV, PASSPORT, TELEPHONE, LDOK, LDATEBEG, LDATEEND, NAPARTMENT, DATEIN, DATEOUT, SROKIN, FLOOR, SUBFAM, SUBCOM, SUBDOK, COMM from " & s & " Where trim(TABN) <> null"
'�� ������� ��������� ���� com �������� ��� ���� ������������ �� �� ���� �������
Cn.Execute "Insert Into data\All_k.dbf (TABN, NUMHABIT, FAM, IM, OT, NLODGER, NLODGERF, NROOM, COMSPACE, HABSPACE, PRIVILEGE, HABITATE, BIRTHDAY, NORDER, DATAORDER, ORGORDER, COSTBUILD, COSTFULL, KITCHSPACE, BATHSPACE, CORRSPACE, TOILSPACE, BALCSPACE, NFAMILY, DATARECEIV, PASSPORT, TELEPHONE, LDOK, LDATEBEG, LDATEEND, NAPARTMENT, DATEIN, DATEOUT, SROKIN, FLOOR, SUBFAM, SUBCOM, SUBDOK) Select TABN, NUMHABIT, FAM, IM, OT, NLODGER, NLODGERF, NROOM, COMSPACE, HABSPACE, PRIVILEGE, HABITATE, BIRTHDAY, NORDER, DATAORDER, ORGORDER, COSTBUILD, COSTFULL, KITCHSPACE, BATHSPACE, CORRSPACE, TOILSPACE, BALCSPACE, NFAMILY, DATARECEIV, PASSPORT, TELEPHONE, LDOK, LDATEBEG, LDATEEND, NAPARTMENT, DATEIN, DATEOUT, SROKIN, FLOOR, SUBFAM, SUBCOM, SUBDOK from " & s & " Where trim(TABN) <> null"
Cn.Execute "UPDATE " & s & " SET BALCSPACE = 0"
Cn.Execute "UPDATE data\all_k SET all_k.DOM = '" & Trim(Str(n)) & "' WHERE(all_k.BALCSPACE < " & n & ")"
Cn.Execute "UPDATE data\all_k SET all_k.BALCSPACE = " & (n + 10)
Cn.Close
'*****************************
Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=����� dBASE;Initial Catalog=data\"
Cn.Execute "UPDATE data\all_k SET all_k.DOM = '" & Right(s, Len(s) - 1) & "' WHERE(all_k.DOM=('" & Trim(Str(n)) & "'))"
Cn.Close
End If






s = Dir()
If Len(s) = 0 Then
Mass.Label2 = "���������� ������ ���������. ��� ������� ������ <�����>, ��� ������, ����� ����������. �������������� ����� ������ �������� �� �������������"
Mass.Label2 = ""
Mass.Label2.Refresh


End If
Wend
Mass.Label2.Caption = "����� ������ �� ���������� <Infin> �������� �������. ����� ���������� � ���� �2"
Mass.Command3.Visible = True
 '*********************************************
Exit Sub
Er:
Select Case Err.Number
'Case Is = 3021
'MsgBox ("��� ����������. �� �������� ��������� ���������� ���������� ���������� (F<3>), ������� ������ �������������� ��� ������� ���������������� ��������� (�� ������ � �����)!")
'��������
Case Is = 0
Case Else
MsgBox (Err.Description)
End Select
End Sub

Private Sub Command5_Click()
'*******************************
   MainMenu.Show
   
 '********************************
End Sub

Private Sub Command6_Click()
Mass.Show
Mass.Label2 = "��������!! ��� �2, ����������� ������ ����� ��������� ���������� ���� �1. ��� ������ ������� ������ ������� <�����>.����� ������� <�����>, ��� ������ ������ ����� ���������� � �� �� ����� �������� �����"

End Sub

Private Sub Command7_Click()
Mass.Show
'Mass.Command4 = True
End Sub

Private Sub Form_Load()
Pt = Pt1
Call MakePt(Pt)

End Sub

Private Sub Text1_Change()
Call MakePt(Pt)
End Sub

Private Sub ������_Click(Index As Integer)
Label1.Caption = "������"
M = "08"
Call MakePt(Pt)
End Sub

Private Sub ������_Click(Index As Integer)
Label1.Caption = "������"
M = "04"
Call MakePt(Pt)
End Sub

Private Sub �����_Click(Index As Integer)
Form1.Hide
Dialog.Hide

End Sub

Private Sub �������_Click(Index As Integer)
Label1.Caption = "�������"
M = "12"
Call MakePt(Pt)
End Sub

Private Sub ����_Click(Index As Integer)
Label1.Caption = "����"
M = "07"
Call MakePt(Pt)
End Sub

Private Sub ����_Click(Index As Integer)
Label1.Caption = "����"
M = "06"
Call MakePt(Pt)
End Sub

Private Sub ���_Click(Index As Integer)
Label1.Caption = "���"
M = "05"
Call MakePt(Pt)
End Sub

Private Sub ����_Click(Index As Integer)
Label1.Caption = "����"
M = "03"
Call MakePt(Pt)
End Sub

Private Sub �����_��������_Click(Index As Integer)
Form3.Show
End Sub

Private Sub ������_Click(Index As Integer)
M = "11"
Call MakePt(Pt)
Label2.Caption = "������"

End Sub

Private Sub �������_Click(Index As Integer)
Label1.Caption = "�������"
M = "10"
Call MakePt(Pt)
End Sub

Private Sub ����������_Click(Index As Integer)
Dim AboutBox As New AboutBox
With AboutBox
    .Title = " ������ � ������ ������������ �������� ���������"
    .Version = "������ 1.0"
    .Company = ""
    .Copyright = " ������� ������ ������������ 2004 ���������"
    .Description = "����������� ������������� ��������"
    .License = "����� � ������� E-Mail:bestonline@list.ru ��������:+79881733600"
    .hWndOwner = Me.hwnd
    'Set .Icon = Me.Icon
    .AboutBox
End With
About.Show
End Sub

Private Sub ����_Click(Index As Integer)
'********************************************
Dialog.Show

End Sub

Private Sub ����������_������_Click(Index As Integer)
Razr.Show

End Sub

Private Sub ��������_Click(Index As Integer)
Label1.Caption = "��������"
M = "09"
Call MakePt(Pt)
End Sub

Private Sub �������_Click(Index As Integer)
Label1.Caption = "�������"
M = "02"
Call MakePt(Pt)
End Sub

Private Sub ������_Click(Index As Integer)
Label1.Caption = "������"
M = "01"
Call MakePt(Pt)
End Sub


Sub ProcessingDBF(ByRef s As String)

'conn.Execute "alter table " & s & " add column Bud char (10)"


End Sub

Sub MakePt(Pt)
God = Form1.Text1
P = Dialog.Dir1.Path
Pt = P + "\G" + God + "\M" + M + "\"
Label2.Caption = Pt

End Sub




