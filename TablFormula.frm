VERSION 5.00
Begin VB.Form TablFormula 
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   1575
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���������� ������� ����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "TablFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Im As String
Dim Dn As String


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Combo1.Text = "��� �������"
Combo1.AddItem "������ �� ������� ��������� �������"
Combo1.AddItem "��������� �� ������� ��������� �������"
Combo1.AddItem "�������� �� ������� ��������� �������"
Combo1.AddItem "������ ���. �� ������� ��������� �������"
Combo1.AddItem "������ ���. �� ������� ��������� �������"
Combo1.AddItem "������ ���. �� ���/��"
Combo1.AddItem "������ ���. �� ���/��"
Combo1.AddItem "�������� �����. �� ���/��"
Combo1.AddItem "��������� �����. �� ���/��"
Combo1.AddItem "�������� �����. �� ���/��"
Combo1.AddItem "�����"
Combo1.AddItem "����������"
Combo1.AddItem "����� �������"
Combo1.AddItem "�������� �������"
Combo1.AddItem "���������"
Combo1.AddItem "���������"
Combo1.AddItem "���-�� ��� ������� ����� �����"
Combo1.AddItem "����"



Combo1.Visible = True
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Combo1.Text <> "" Then
Im = ""
If DocTBL.Naz = "S1" Then Im = "F1"
If DocTBL.Naz = "S2" Then Im = "F2"
If DocTBL.Naz = "S3" Then Im = "F3"
If DocTBL.Naz = "S4" Then Im = "F4"
If DocTBL.Naz = "S5" Then Im = "F5"

Dn = Combo1.Text
         If Im <> "" Then
Dn = Chr(34) + Dn + Chr(34)
Q = "UPDATE TablDoc SET TablDoc." + Im + " = " + Dn + " WHERE (((TablDoc.Cod)=" + DocTBL.L + "))"
'MsgBox Q
mconn.Execute (Q)
          End If

Unload Me
Else
MsgBox "������ ��������"
End If
End Sub

Private Sub Command5_Click()
Unload TablFormula
End Sub

Private Sub Form_Unload(Cancel As Integer)
DocTBL.Enabled = True
End Sub
