VERSION 5.00
Begin VB.Form RepSverka 
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   7
      Text            =   "���"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Text            =   "���"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   5295
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�� ��� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�� ����� �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   1920
      Width           =   5295
   End
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����������������?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��������������?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "����� ���������� ������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4050
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
      Height          =   195
      Left            =   5280
      Picture         =   "RepSverka.frx":0000
      Top             =   720
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "RepSverka.frx":024A
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   5160
      Picture         =   "RepSverka.frx":0994
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4800
      Picture         =   "RepSverka.frx":10DE
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "RepSverka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Addrconn As ADODB.Recordset
'Dim mconn As ADODB.Connection

Private Sub BtnEnh1_Click()
MainMenu.Enabled = True
Unload Me
End Sub

Private Sub BtnEnh2_1_Click()

End Sub

Private Sub BtnEnh2_Click()
' �� ���.������
Dim sq As String
Dim Sort As String
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))
filA = Val(Replace(Combo2.Text, " ", "_", 1))


If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" Then StrU = ""
If Combo1.Text <> "��� ����" And Combo2.Text = "��� ����������" Then StrU = "WHERE (((KLS_PODR.���)=" + Str(fil) + "))"
If Combo1.Text <> "��� ����" And Combo2.Text <> "��� ����������" Then StrU = "WHERE (((KLS_PODR.���)=" + Str(fil) + " And (Adding.KodN=" + Str(filA) + ")))"
If Combo1.Text = "��� ����" And Combo2.Text <> "��� ����������" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + "))"


If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text <> "���" And Combo4.Text = "���" Then StrU = "WHERE (((KLS_PODR.����)='" + Combo3.Text + "'))"
If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((MainOccupant.priv)='" + Combo4.Text + "'))"
If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text <> "���" And Combo4.Text <> "���" Then StrU = "WHERE (((KLS_PODR.����)='" + Combo3.Text + "') AND ((MainOccupant.Priv)='" + Combo4.Text + "'))"

If Combo1.Text <> "��� ����" And Combo3.Text <> "���" Then
MsgBox "������� ������ ��������� �������." + vbNewLine + " ���� �� ������ ������� o���� �� ���������������� �����, �� ���� ������� <��� ����>"
Exit Sub
End If


If Combo1.Text <> "��� ����" And Combo2.Text <> "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + ") AND ((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.���)=" + Str(fil) + "))"
If Combo1.Text <> "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.���)=" + Str(fil) + "))"
If Combo1.Text = "��� ����" And Combo2.Text <> "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((Adding.KodN)=" + Str(filA) + "))"



Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " �., �� ������:" + Combo1.Text
Analizlgot.G = 9
'sq = "SELECT KLS_PODR.NAIM_KLS as �����,  MainOccupant.KV_NUM as ��,MainOccupant.FAM as �������, MainOccupant.IM as ���, MainOccupant.OT as ��������,Adding.KodN as ���,Adding.NameN as ����������,   Adding.SummaI FROM Adding INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) ON Adding.KodKv = MainOccupant.Numer " + StrU
sq = "SELECT KLS_PODR.NAIM_KLS as �����,  MainOccupant.KV_NUM as ��,MainOccupant.FAM as �������, MainOccupant.IM as ���, MainOccupant.OT as ��������,Adding.KodN as ���,Adding.NameN as ����������,   IIf([Adding]![Tip]='+',[SummaI],0) AS ���������, IIf([Adding]![Tip]='-',[SummaI],0) AS ��������, IIf([Adding]![Tip]='s',[SummaI],0) AS �������� FROM Adding INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) ON Adding.KodKv = MainOccupant.Numer " + StrU
Analizlgot.G = 11
Analizlgot.StrSQL = sq
Analizlgot.Show
Analizlgot.FG1.AutoResize = True
Unload Me
Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
Analizlgot.FG1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
End Sub

Private Sub BtnEnh3_Click()
'�� ����� �������

Dim sq As String
Dim Sort As String
Dim StrU As String
Dim fil As Integer
Dim filA As Integer

fil = Val(Replace(Combo1.Text, " ", "_", 1))
filA = Val(Replace(Combo2.Text, " ", "_", 1))

If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text = "���" And Combo4.Text = "���" Then StrU = ""
If Combo1.Text <> "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text = "���" And Combo4.Text = "���" Then StrU = "WHERE (((KLS_PODR.���)=" + Str(fil) + "))"
If Combo1.Text <> "��� ����" And Combo2.Text <> "��� ����������" And Combo3.Text = "���" And Combo4.Text = "���" Then StrU = "WHERE (((KLS_PODR.���)=" + Str(fil) + " And (Adding.KodN=" + Str(filA) + ")))"
If Combo1.Text = "��� ����" And Combo2.Text <> "��� ����������" And Combo3.Text = "���" And Combo4.Text = "���" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + "))"

If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text <> "���" And Combo4.Text = "���" Then StrU = "WHERE (((KLS_PODR.����)='" + Combo3.Text + "'))"
If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((MainOccupant.priv)='" + Combo4.Text + "'))"
If Combo1.Text = "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text <> "���" And Combo4.Text <> "���" Then StrU = "WHERE (((KLS_PODR.����)='" + Combo3.Text + "') AND ((MainOccupant.Priv)='" + Combo4.Text + "'))"

If Combo1.Text <> "��� ����" And Combo3.Text <> "���" Then
MsgBox "������� ������ ��������� �������." + vbNewLine + " ���� �� ������ ������� o���� �� ���������������� �����, �� ���� ������� <��� ����>"
Exit Sub
End If


If Combo1.Text <> "��� ����" And Combo2.Text <> "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + ") AND ((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.���)=" + Str(fil) + "))"
If Combo1.Text <> "��� ����" And Combo2.Text = "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.���)=" + Str(fil) + "))"
If Combo1.Text = "��� ����" And Combo2.Text <> "��� ����������" And Combo3.Text = "���" And Combo4.Text <> "���" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((Adding.KodN)=" + Str(filA) + "))"


Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " �., �� ������:" + Combo1.Text
Analizlgot.G = 9
sq = "SELECT KLS_PODR.NAIM_KLS as �����, Adding.KodN as ���, Adding.NameN as ����������, MainOccupant.KV_NUM as ��, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.SummaI FROM Adding INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) ON Adding.KodKv = MainOccupant.Numer " + StrU
Analizlgot.G = 9
Analizlgot.StrSQL = sq
Analizlgot.Show
Analizlgot.FG1.AutoResize = True
Unload Me
Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"



End Sub


Private Sub Form_Load()
MakeWindow Me, True

  'Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"






Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic


Set Nconn = New ADODB.Recordset
Set Nconn.ActiveConnection = Mconn
Nconn.CursorType = adOpenStatic
Nconn.LockType = adLockBatchOptimistic



'AddrConn.Open ("KLS_PODR")
Addrconn.Open ("SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.�������������, KLS_PODR.���� From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Combo1.Text = "��� ����"

'��� ���������� �������
Addrconn.MoveFirst
Combo1.AddItem "��� ����"
Do While Not Addrconn.EOF
If Addrconn("���") <> -1 Then
Combo1.AddItem Trim(Str(Addrconn("���"))) + " " + Addrconn("NAIM_KLS") + " ��� � " + Addrconn("Num")
End If
Addrconn.MoveNext
Loop

'��� ���������� ����������
Nconn.Open ("SELECT nachisleniy.Kod, nachisleniy.Naim From Nachisleniy ORDER BY nachisleniy.Kod DESC")
Combo2.Text = "��� ����������"
Nconn.MoveFirst
Combo2.AddItem "��� ����������"
Do While Not Nconn.EOF
Combo2.AddItem Trim(Str(Nconn("kod"))) + " " + Nconn("NAIM")
Nconn.MoveNext
Loop
Nconn.Close
Addrconn.Close

Combo3.AddItem "���"
Combo3.AddItem "���������."
Combo3.AddItem "�����������."


Combo4.AddItem "���"
Combo4.AddItem "��"
Combo4.AddItem "���"

End Sub


