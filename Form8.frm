VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   0  'None
   ClientHeight    =   3312
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5616
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   5295
      _ExtentX        =   9335
      _ExtentY        =   868
      Caption         =   "������ �����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "������"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton BtnEnh4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "� ������� �����"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "���������"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�����������"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�� ���.�������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   $"Form8.frx":0000
      Top             =   2640
      Width           =   3255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�� ���-�� ������."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   $"Form8.frx":009C
      Top             =   2400
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�� � ��������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   $"Form8.frx":0138
      Top             =   2160
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�� �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   $"Form8.frx":01D4
      Top             =   1920
      Width           =   3255
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
      ToolTipText     =   "������ ������ ����� ������� �� ���������� ������"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "�����������"
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
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "����� ���������� ������"
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
      Height          =   156
      Left            =   5280
      Picture         =   "Form8.frx":0270
      Top             =   720
      Width           =   156
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   4800
      Picture         =   "Form8.frx":04BA
      Top             =   120
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   5160
      Picture         =   "Form8.frx":0C04
      Top             =   120
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4800
      Picture         =   "Form8.frx":134E
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Addrconn As ADODB.Recordset
Dim fil As Integer

'Dim mconn As ADODB.Connection

Private Sub BtnEnh1_Click()
MainMenu.Enabled = True
Unload Me
End Sub

Private Sub BtnEnh2_1_Click()

End Sub

Private Sub BtnEnh2_Click()
Dim sq As String
Dim Sort As String


If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If

fil = Val(Replace(Combo1.Text, " ", "_", 1))




If Option1.Value = True Then Sort = "ORDER BY MainOccupant.FAM"
If Option2.Value = True Then Sort = "ORDER BY MainOccupant.kv_num"
If Option3.Value = True Then Sort = "ORDER BY MainOccupant.NLODGERF"
If Option4.Value = True Then Sort = "ORDER BY MainOccupant.COMSPACE"



sq = "SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.bankN as N,MainOccupant.FAM as �������, MainOccupant.IM as ���, MainOccupant.OT as ��������, MainOccupant.kv_num as [�� �], MainOccupant.COMSPACE as [����� ��], MainOccupant.NLODGERF as ���������, Sum((Adding!SaldoN*1000/Adding!Kol)/1000) AS [������� ���], Sum(IIf(Adding!Tip=" + Chr(34) + "+" + Chr(34) + ",[SummaI],0)) AS ���������, Sum(IIf(Adding!Tip=" + Chr(34) + "s" + Chr(34) + ",[SummaI],0)) AS ��������, Sum(IIf(Adding!Tip=" + Chr(34) + "-" + Chr(34) + ",[SummaI],0)) AS ������, Sum((Adding!SaldoK*1000/Adding!Kol)/1000) AS [������� ���] FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.bankN ,MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.COMSPACE, MainOccupant.NLODGERF Having (((KLS_PODR.���) =" + Str(fil) + "))" + Sort

Analizlgot.G = 16



If MainForm.Dog = 1 Then
 sq = "SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.bankN as N,MainOccupant.FAM as �������, MainOccupant.IM as ���, MainOccupant.OT as ��������, MainOccupant.kv_num as [�� �], MainOccupant.COMSPACE as [����� ��], MainOccupant.NLODGERF as ���������, Sum((Adding!SaldoN*1000/Adding!Kol)/1000) AS [������� ���], Sum(IIf(Adding!Tip=" + Chr(34) + "+" + Chr(34) + ",[SummaI],0)) AS ���������, Sum(IIf(Adding!Tip=" + Chr(34) + "s" + Chr(34) + ",[SummaI],0)) AS ��������, Sum(IIf(Adding!Tip=" + Chr(34) + "-" + Chr(34) + ",[SummaI],0)) AS ������, Sum((Adding!SaldoK*1000/Adding!Kol)/1000) AS [������� ���], MainOccupant.Dog as [���� ������]"
 sq = sq + " FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.bankN ,MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.COMSPACE, MainOccupant.NLODGERF, MainOccupant.Dog Having (((KLS_PODR.���) =" + Str(fil) + "))" + Sort
 Analizlgot.G = 17
 'Analizlgot.FG1.Subtotal flexSTSum, 1, 1, , RGB(150, 250, 200), vbBlack, True
End If
Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " �., �� ������:" + Combo1.Text



Analizlgot.StrSQL = sq
Analizlgot.Show



Analizlgot.fg1.ColHidden(1) = True
Analizlgot.fg1.ColHidden(2) = True
Analizlgot.fg1.ColHidden(3) = True

Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
Analizlgot.fg1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True


If MainForm.Dog = 1 Then Analizlgot.fg1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True

'Analizlgot.FG1.Subtotal flexSTSum, 4, 10, , RGB(250, 250, 200), vbBlack, True, "� ���� �/��:"
'Analizlgot.FG1.Subtotal flexSTSum, 4, 11, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 12, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 13, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 14, , RGB(250, 250, 200), vbBlack, True





Unload Me
'Analizlgot.�� 1

End Sub

Private Sub BtnEnh3_Click()

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub

End If


Dim sq As String
Dim Sort As String

Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

If Option1.Value = True Then Sort = "ORDER BY MainOccupant.FAM"
If Option2.Value = True Then Sort = "ORDER BY MainOccupant.kv_num"
If Option3.Value = True Then Sort = "ORDER BY MainOccupant.NLODGERF"
If Option4.Value = True Then Sort = "ORDER BY MainOccupant.COMSPACE"




sq = "SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������,"
sq = sq + "MainOccupant.kv_num AS [�� �], MainOccupant.COMSPACE AS [����� ��], MainOccupant.NLODGERF AS ���������,  Adding.NameKat as [��������� �������],Sum((Adding!SaldoN*1000/Adding!Kol)/1000) AS [������� ���], Sum(IIf(Adding!Tip=" + Chr(34) + "+" + Chr(34) + ",[SummaI],0)) AS ���������, Sum(IIf(Adding!Tip=" + Chr(34) + "s" + Chr(34) + ",[SummaI],0)) AS ��������, Sum(IIf(Adding!Tip=" + Chr(34) + "-" + Chr(34) + ",[SummaI],0)) AS ������, Sum((Adding!SaldoK*1000/Adding!Kol)/1000) AS [������� ���] FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, Adding.NameKat, MainOccupant.COMSPACE, MainOccupant.NLODGERF HAVING (((KLS_PODR.���)=" + Str(fil) + "))" + Sort


Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " �., �� ������:" + Combo1.Text


Analizlgot.G = 16
Analizlgot.StrSQL = sq
Analizlgot.Show

Analizlgot.fg1.ColHidden(1) = True
Analizlgot.fg1.ColHidden(2) = True
Analizlgot.fg1.ColHidden(3) = True

Analizlgot.fg1.AutoResize = True



Unload Me
'Analizlgot.�� 1
Analizlgot.fg1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
Analizlgot.fg1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True

Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True

Analizlgot.fg1.Subtotal flexSTSum, 4, 11, , RGB(250, 250, 200), vbBlack, True, "� ���� �/��:"
Analizlgot.fg1.Subtotal flexSTSum, 4, 12, , RGB(250, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 4, 13, , RGB(250, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 4, 14, , RGB(250, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 4, 15, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 12, , RGB(250, 250, 200), vbBlack, True, "����:"
End Sub


Private Sub BtnEnh4_Click()

Dim sq As String
Dim fil As Integer
Dim Sort As String

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If

fil = Val(Replace(Combo1.Text, " ", "_", 1))




If Option1.Value = True Then Sort = "ORDER BY MainOccupant.FAM"
If Option2.Value = True Then Sort = "ORDER BY MainOccupant.kv_num"
If Option3.Value = True Then Sort = "ORDER BY MainOccupant.NLODGERF"
If Option4.Value = True Then Sort = "ORDER BY MainOccupant.COMSPACE"


'sq = "SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.bankN as N,MainOccupant.FAM as �������, MainOccupant.IM as ���, MainOccupant.OT as ��������, MainOccupant.kv_num as [�� �], MainOccupant.COMSPACE as [����� ��], MainOccupant.NLODGERF as ���������, Sum((Adding!SaldoN*1000/Adding!Kol)/1000) AS [������� ���], Sum(IIf(Adding!Tip=" + Chr(34) + "+" + Chr(34) + ",[SummaI],0)) AS ���������, Sum(IIf(Adding!Tip=" + Chr(34) + "s" + Chr(34) + ",[SummaI],0)) AS ��������, Sum(IIf(Adding!Tip=" + Chr(34) + "-" + Chr(34) + ",[SummaI],0)) AS ������, Sum((Adding!SaldoK*1000/Adding!Kol)/1000) AS [������� ���] FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.bankN ,MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.COMSPACE, MainOccupant.NLODGERF Having (((KLS_PODR.���) =" + Str(fil) + "))" + Sort

sq = "SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.BanKN AS N, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.kv_num AS [�� �], MainOccupant.COMSPACE AS [����� ��], MainOccupant.NLODGERF AS ���������, Sum(([Adding]![SaldoN]*1000/[Adding]![Kol])/1000) AS [������� ���], Sum(IIf([Adding]![Tip]='+',[SummaI],0)) AS ���������, Sum(IIf([Adding]![Tip]='s',[SummaI],0)) AS ��������, Sum(IIf([Adding]![Tip]='-',[SummaI],0)) AS ������, Sum(([Adding]![SaldoK]*1000/[Adding]![Kol])/1000) AS [������� ���], Lgota.Numer AS ���, Lgota.NAME_KLS AS ������, IIf([Lgota]![OhteCode] is not null,IIf([Lgota]![OhteCode]<>0,'����.������','���.��/�����. '),null) AS [�������������� ������] FROM Lgota RIGHT JOIN ((Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) ON Lgota.NomNum = MainOccupant.Numer "



sq = sq + "GROUP BY KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.BanKN, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.COMSPACE, MainOccupant.NLODGERF, Lgota.Numer, Lgota.NAME_KLS, IIf([Lgota]![OhteCode] is not null,IIf([Lgota]![OhteCode]<>0,'����.������','���.��/�����. '),null) HAVING (((KLS_PODR.���)= " + Str(fil) + "))" + Sort

Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " �., �� ������:" + Combo1.Text


Analizlgot.G = 19
Analizlgot.StrSQL = sq
'MsgBox sq
Analizlgot.Show



Analizlgot.fg1.ColHidden(1) = True
Analizlgot.fg1.ColHidden(2) = True
Analizlgot.fg1.ColHidden(3) = True

Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
Analizlgot.fg1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True

'Analizlgot.FG1.Subtotal flexSTSum, 4, 10, , RGB(250, 250, 200), vbBlack, True, "� ���� �/��:"
'Analizlgot.FG1.Subtotal flexSTSum, 4, 11, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 12, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 13, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 14, , RGB(250, 250, 200), vbBlack, True





Unload Me
'Analizlgot.�� 1

End Sub





Private Sub Check1_Click()
If Check1.Value = 1 Then Check2.Value = 0
If Check1.Value = 0 Then Check2.Value = 1
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Check1.Value = 0
If Check2.Value = 0 Then Check1.Value = 1
End Sub

Private Sub Form_Load()
MakeWindow Me, True

Option1.Value = True


'Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"



Option1.BackColor = RGB(207, 207, 207)
Option2.BackColor = RGB(207, 207, 207)
Option3.BackColor = RGB(207, 207, 207)
Option4.BackColor = RGB(207, 207, 207)

Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic


'AddrConn.Open ("KLS_PODR")
Addrconn.Open ("SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.�������������, KLS_PODR.���� From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Combo1.Text = "������ �����"


Addrconn.MoveFirst
Combo1.AddItem "��� ����"
Do While Not Addrconn.EOF
If Addrconn("���") <> -1 Then
Combo1.AddItem Trim(Str(Addrconn("���"))) + " " + Addrconn("NAIM_KLS") + " ��� � " + Addrconn("Num")
End If
Addrconn.MoveNext
Loop




End Sub


Private Function Addres(KLS As String) As String


End Function

Private Sub xpcmdbutton1_Click()
VibDom.Show
End Sub
