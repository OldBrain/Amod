VERSION 5.00
Begin VB.Form RepParam1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7740
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "RepParam1.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   477
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����� � ������������� �����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   6975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������ �/��. � ���������� �������� ������� ����� � ������������ ��������, '����� ������'. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "������ �������� �������� �� ��� �����"
      Top             =   6600
      Width           =   6975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������ �/��. � ���������� �������� ������� ����� � ������������ ��������, '����������'. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "������ �������� �������� �� ��� �����"
      Top             =   5760
      Width           =   6975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����� � ������������� ���������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   6975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ����� ����� ����������� ������������ �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ����� ����� ��������� ������������ �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ���.������"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   3480
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
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "���"
      Top             =   1920
      Width           =   5775
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2520
      TabIndex        =   4
      Text            =   "���"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����������"
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
      TabIndex        =   1
      Top             =   2760
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2040
      TabIndex        =   0
      Text            =   "���"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "� ����� �� ���������� ""�������"" ������������ ����� "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��� ""���������� + "" ������ � ������"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4170
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
      Height          =   240
      Left            =   0
      Picture         =   "RepParam1.frx":030A
      ToolTipText     =   "�������"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   960
      Picture         =   "RepParam1.frx":084C
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   360
      Picture         =   "RepParam1.frx":0F96
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   2040
      Picture         =   "RepParam1.frx":16E0
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
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
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��������� �������:"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "������ ��:"
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
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "RepParam1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ��������� As String
Dim ������������� As String
Dim ����� As String

Private Sub Combo1_Validate(Cancel As Boolean)
������������� = Trim(Combo1.Text)

End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
��������� = Combo2.Text
End Sub

Private Sub Combo3_Click()
If Combo3.Text <> "���" Then
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False

Else
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
End If
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
����� = Combo3.Text


End Sub

Private Sub Command1_Click()

If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "���������� �������" And ����� = "���" Then

Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18

Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [��� �����]-[���������] AS [� ����������], tmp_lgota.NAME_KLS AS ������������, Sum([tmp_lgota]![Prim1]) AS [���������� ��� �������], tmp_lgota.Procent AS [������� �����], tmp_lgota.Use, [Adding]![Tarif]*[���������� ��� �������]*[tmp_lgota]![Procent]/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom"
Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"


'Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], Sum([��� �����]-[���������]) AS [� ����������], tmp_lgota.NAME_KLS AS ������������, tmp_lgota.PloLG AS [��� ���], tmp_lgota.Procent AS [������� �����], [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom"
'Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.PloLG, tmp_lgota.Procent, [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS"


Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 150, 200), vbBlack, True, "�����"

Analizlgot.FG1.Subtotal flexSTSum, 0, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 17, , RGB(150, 250, 200), vbBlack, True



Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"

Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, True




End If


If Combo1.Text = "���������� �������" And ����� <> "���" Then

Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18

Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [��� �����]-[���������] AS [� ����������], tmp_lgota.NAME_KLS AS ������������, Sum([tmp_lgota]![Prim1]) AS [���������� ��� �������], tmp_lgota.Procent AS [������� �����], tmp_lgota.Use, [Adding]![Tarif]*[���������� ��� �������]*[tmp_lgota]![Procent]/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom"
Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((KLS_PODR.NAIM_KLS)='" + ����� + "') AND ((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"



'Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], Sum([��� �����]-[���������]) AS [� ����������], tmp_lgota.NAME_KLS AS ������������, tmp_lgota.PloLG AS [��� ���], tmp_lgota.Procent AS [������� �����], [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom"
'Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.PloLG, tmp_lgota.Procent, [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS"


'Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 250, 200), vbBlack, True, "�����"

Analizlgot.FG1.Subtotal flexSTSum, 0, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 17, , RGB(150, 250, 200), vbBlack, True



Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"

Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, True




End If




If Combo1.Text = "�������" And ����� = "���" Then
Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [��� �����]-[���������] AS [� ����������], tmp_lgota.NAME_KLS AS ������������, Sum(tmp_lgota.PloLG) AS [��� ���], tmp_lgota.Procent AS [������� �����], Adding!Tarif*[��� ���]*tmp_lgota!Procent/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key=tmp_lgota.UniKOd) ON MainOccupant.Numer=Adding.KodKv) ON KLS_PODR.���=MainOccupant.Dom"
Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"



Analizlgot.FG1.Subtotal flexSTSum, 0, 14, , RGB(150, 250, 200), vbBlack, True, "�����"
'Analizlgot.FG1.Subtotal flexSTSum, 0, 15, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 17, , RGB(150, 250, 200), vbBlack, True



Analizlgot.FG1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
'Analizlgot.FG1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, True




End If

If Combo1.Text = "�������" And ����� <> "���" Then

Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [��� �����]-[���������] AS [� ����������], tmp_lgota.NAME_KLS AS ������������, Sum(tmp_lgota.PloLG) AS [��� ���], tmp_lgota.Procent AS [������� �����], Adding!Tarif*[��� ���]*tmp_lgota!Procent/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.ispr, Adding.NameKat, tmp_lgota.Prim"


Reports.sq = Reports.sq + " HAVING (((KLS_PODR.NAIM_KLS)='" + ����� + "') AND ((Adding.ispr)=0) AND ((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"


Analizlgot.FG1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True, "� ���� �� ����"
'Analizlgot.FG1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, t

End If



'MsgBox Reports.sq
'Analizlgot.�� 2
Analizlgot.Show

Unload Me
'Unload RepLgota
Unload Reports
Exit Sub
'Unload RepLgota
Unload Me
End Sub

Private Sub Command2_Click()
If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "�������" And ����� = "���" Then
Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 14
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, KLS_PODR.Tip_Naim, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [Adding]![SummaBl]-[Adding]![SummaI] AS [� ����������], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN Adding ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom GROUP BY KLS_PODR.NAIM_KLS, KLS_PODR.Tip_Naim, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat Having ((([Adding]![SummaBl] - [Adding]![SummaI]) <> 0) And ((Adding.ispr) = 0) And ((Adding.NameKat) = " + Chr(34) + Combo2.Text + Chr(34) + ")) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"

'Reports.sq = Reports.sq + ""

Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 250, 200), vbBlack, True, "�����"
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 12, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True


End If

If Combo1.Text = "�������" And ����� <> "���" Then

Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 14
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, KLS_PODR.Tip_Naim, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [��� �����]-[���������] AS [� ����������], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom GROUP BY KLS_PODR.NAIM_KLS, KLS_PODR.Tip_Naim, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((KLS_PODR.NAIM_KLS)='" + ����� + "') AND ((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"
'Reports.sq = Reports.sq + ""


Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 250, 200), vbBlack, True, "�����"
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 12, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True


End If




Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
Exit Sub

Unload Me

End Sub

Private Sub Command3_Click()
If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "�������" Then
Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������� �����], tmp_lgota.Use, Adding.Tarif AS �����, round(Sum(tmp_lgota.PloLG),2) AS [��� ���], round(([Adding]![Tarif]*[��� ���]*[tmp_lgota]![Procent]/100),2) AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS, Sum(tmp_lgota.PloLG)"


Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True, "�����"
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True

'Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True


End If

If Combo1.Text = "���������� �������" Then
Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8

Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������� �����], tmp_lgota.Use, Adding.Tarif AS �����, Sum([tmp_lgota]![Prim1]) AS [��� �� ��� �������], Round(Sum(([Adding]![Tarif]*[tmp_lgota]![Prim1]*[tmp_lgota]![Procent]/100)),2) AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS"

'Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True, "� ����:"
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True

End If


Analizlgot.Show
Unload Me
Unload RepLgota
Unload Reports
Exit Sub

Unload Me

End Sub

Private Sub Command4_Click()
If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "�������" Then
Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8

Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������� �����], tmp_lgota.Use, Adding.Tarif AS �����, round(tmp_lgota.PloLG,2) AS [��� ���], Round(Sum(([Adding]![Tarif]*[��� ���]*[tmp_lgota]![Procent]/100)),2) AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, Adding.ispr, tmp_lgota.PloLG, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS, tmp_lgota.PloLG"

Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True

End If


If Combo1.Text = "���������� �������" Then
Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8

Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������� �����], tmp_lgota.Use, Adding.Tarif AS �����, [tmp_lgota]![Prim1] AS [��� �� ��� �������], Round(Sum(([Adding]![Tarif]*[tmp_lgota]![Prim1]*[tmp_lgota]![Procent]/100)),2) AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, [tmp_lgota]![Prim1], Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS"

Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True

End If





Analizlgot.Show
Unload Me
Unload RepLgota
Unload Reports
End Sub

Private Sub Command5_Click()
If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True


Exit Sub
End If

Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 10
'Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������ �����], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.Propis AS [���-�� �� ���], Sum(tmp_lgota.PloLG) AS [��� �������], Adding.ObPl AS [��� ��], Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], [��� �����]-[���������] AS [� ����������] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1))"

'Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������ �����], Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.Propis AS [���-�� �� ���], Sum(tmp_lgota.PloLG) AS [��� �������], Adding.ObPl AS [��� ��], Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], ([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100 AS [� ������������] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, ([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"

If Combo1.Text = "�������" Then
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������ �����], Adding.Propis AS [���-�� �� ���], Sum(tmp_lgota.PloLG) AS [��� �������], Adding.ObPl AS [��� ��], Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], Sum(([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100) AS [� ������������],  Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"
End If


If Combo1.Text = "���������� �������" Then
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS ������������, tmp_lgota.Procent AS [������ �����], Adding.Propis AS [���-�� �� ���], Sum(tmp_lgota.Prim1) AS [��� ��� �������], Adding.ObPl AS [��� ��], Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], Sum(([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![Prim1])/100) AS [� ������������],  Count(tmp_lgota.UniKOd) AS [���-�� �����], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"
End If



Analizlgot.�� 2

'Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTClear, 2, 7, , RGB(150, 250, 100), vbBlack, True

Analizlgot.Show




Unload Me
Unload RepLgota
Unload Reports
Exit Sub
'Unload RepLgota


End Sub

Private Sub Command6_Click()

If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If



Analizlgot.Titl = "���������� ��������� ����������������� ����� �� " + vbNewLine + "������������ ������� �����, � ��� �� ������������ ��������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 7


If Combo1.Text = "�������" Then
Reports.sq = "SELECT TEST_02.�, TEST_02.�������, TEST_02.���, TEST_02.��������, TEST_02.���������, TEST_02.[��� �����], Test_03.[Sum-� ���-��], TEST_02.[� ����������], Round([TEST_02]![� ����������]-[TEST_03]![Sum-� ���-��],2) AS ����������, TEST_02.NameKat FROM TEST_02 INNER JOIN Test_03 ON TEST_02.� = Test_03.� WHERE (((Round([TEST_02]![� ����������]-[TEST_03]![Sum-� ���-��],2))<-0.01 Or (Round([TEST_02]![� ����������]-[TEST_03]![Sum-� ���-��],2))>0.01) AND ((TEST_02.NameKat)='" + Combo2.Text + "')) ORDER BY Round([TEST_02]![� ����������]-[TEST_03]![Sum-� ���-��],2)"
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 2, , RGB(150, 250, 200), vbBlack, True
End If

'If Combo1.Text = "���������� �������" Then
'Reports.sq = "SELECT [02].�, [02].�������, [02].���, [02].��������, Round([02]![� ����������]-[03]![Sum-� ���-��],2) AS �����������, [02].�����, [02].���������, [02].[��� �����], [02].[� ����������], [02].NameKat FROM 02 INNER JOIN 03 ON [02].� = [03].� Where (((Round([02]![� ����������] - [03]![Sum-� ���-��], 2)) < -0.01 Or (Round([02]![� ����������] - [03]![Sum-� ���-��], 2)) > 0.01) And (([02].NameKat) = '" + Combo2.Text + "')) ORDER BY Round([02]![� ����������]-[03]![Sum-� ���-��],2)"
'End If


'Analizlgot.�� 2
Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
End Sub

Private Sub Command7_Click()
If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If



Analizlgot.Titl = "���������� ��������� ����������������� ����� �� " + vbNewLine + "������������ ������� �����, � ��� �� ������������ ��������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 12


'If Combo1.Text = "�������" Then
Reports.sq = "SELECT TEST_04.�, TEST_04.�������, TEST_04.���, TEST_04.��������, TEST_04.[���������� ��� �������], TEST_04.�����, Round(TEST_04![� ����������]-Test_05![Sum-� ���-��],2) AS �����������, TEST_04.���������, TEST_04.[��� �����], TEST_04.[� ����������], TEST_05.[Sum-� ���-��], TEST_04.NameKat, TEST_04.[���������� ��� �������] FROM TEST_04 INNER JOIN TEST_05 ON TEST_04.� = TEST_05.� WHERE (((TEST_04.NameKat)='����� ������') AND ((Round([TEST_04]![� ����������]-[Test_05]![Sum-� ���-��],2))<-0.01 Or (Round([TEST_04]![� ����������]-[Test_05]![Sum-� ���-��],2))>0.01)) ORDER BY Round(TEST_04![� ����������]-Test_05![Sum-� ���-��],2)"

Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 12, , RGB(150, 250, 200), vbBlack, True
'End If


'Analizlgot.�� 2
Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
End Sub

Private Sub Command8_Click()
If Combo2.Text = "���" Then
MsgBox "������ ���������"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

Analizlgot.Titl = "����� � ����������� ��������, ��������� � ��������� ��� ���������� ��������� " + vbNewLine + "��������� ��������� ������� � ����� ������ �������-������������ ����� �� " + MainForm.Label8 + " �."

If Combo1.Text = "�������" Then

If MsgBox("����������� ������������ ������� �� ������ 18,21 � 33 ��.�. � ������ ?", vbYesNo) = vbYes Then

Analizlgot.G = 12
Reports.sq = "SELECT LGTip.Name AS [��� ������], TMP_Lgota.NAME_KLS AS ������, TMP_Lgota.Use AS [������ ����������], TMP_Lgota.Procent AS [������ �����], TMP_Lgota.tarif, IIf([PloLG]<>18,IIf([PloLG]<>21,IIf([PloLG]<>33,0,TMP_Lgota!PloLG),TMP_Lgota!PloLG),TMP_Lgota!PloLG) AS [������������ �������], Sum(TMP_Lgota.Prop) AS [����� ���������], Count([TMP_Lgota]![Key]) AS ����������, [���������� �����]-[����������] AS [����� �����], Sum(TMP_Lgota!Koll) AS [���������� �����], Round(Sum(([TMP_Lgota]![Procent]*[TMP_Lgota]![PloLG]/100)*[TMP_Lgota]![tarif]),2) AS [� ����������] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif, IIf([PloLG]<>18,IIf([PloLG]<>21,IIf([PloLG]<>33,0,TMP_Lgota!PloLG),TMP_Lgota!PloLG),TMP_Lgota!PloLG)"

Analizlgot.FG1.MergeCells = flexMergeFree



Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 11, , RGB(150, 250, 200), vbBlack, True

Else

Analizlgot.G = 12
Reports.sq = "SELECT LGTip.Name AS [��� ������], TMP_Lgota.NAME_KLS AS ������, TMP_Lgota.Use AS [������ ����������], TMP_Lgota.Procent AS [������ �����], TMP_Lgota.tarif AS �����, [TMP_Lgota]![PloLG] AS [������������ �������], Sum(TMP_Lgota.Prop) AS [����� ���������], Count(TMP_Lgota!Key) AS ����������, [���������� �����]-[����������] AS [����� �����], Sum(TMP_Lgota!Koll) AS [���������� �����], Round(Sum((TMP_Lgota!Procent*TMP_Lgota!PloLG/100)*TMP_Lgota!tarif),2) AS [� ����������] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif, [TMP_Lgota]![PloLG]"

Analizlgot.FG1.MergeCells = flexMergeFree



Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 11, , RGB(150, 250, 200), vbBlack, True



End If


Analizlgot.Show

Unload RepLgota
Unload Reports
Unload Me




End If

If Combo1.Text = "���������� �������" Then


If MsgBox("����������� �� ���������� �����������?", vbYesNo) = vbYes Then


Analizlgot.G = 11
Reports.sq = "SELECT LGTip.Name AS [��� ������], TMP_Lgota.NAME_KLS AS ������, TMP_Lgota.Use AS [������ ����������], TMP_Lgota.Procent AS [������ �����], TMP_Lgota.tarif, TMP_Lgota.Prop AS [����� ���������], Count(TMP_Lgota!Key) AS ����������, [���������� �����]-[����������] AS [����� �����], Sum(TMP_Lgota!Koll) AS [���������� �����], Round(Sum(([TMP_Lgota]![Procent]*[TMP_Lgota]![Koll]/100)*[TMP_Lgota]![tarif]),2) AS [� ����������] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif, TMP_Lgota.Prop"
Analizlgot.FG1.MergeCells = flexMergeFree


Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 2, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 10, , RGB(150, 250, 200), vbBlack, True


Else

Analizlgot.G = 10
Reports.sq = "SELECT LGTip.Name AS [��� ������], TMP_Lgota.NAME_KLS AS ������, TMP_Lgota.Use AS [������ ����������], TMP_Lgota.Procent AS [������ �����], TMP_Lgota.tarif, Count(TMP_Lgota!Key) AS ����������, [���������� �����]-[����������] AS [����� �����], Sum(TMP_Lgota!Koll) AS [���������� �����], Round(Sum((TMP_Lgota!Procent*TMP_Lgota!Koll/100)*TMP_Lgota!tarif),2) AS [� ����������] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif"
Analizlgot.FG1.MergeCells = flexMergeFree


Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 2, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True




End If





Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports



End If

End Sub

Private Sub Form_Load()
Dim cnParam As ADODB.Connection
Dim rsVrem As ADODB.Recordset

MakeWindow Me, True

������� "kvartplata.amd"

'Set cnParam = New ADODB.Connection
 ' cnParam.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' cnParam.Open "data/Kvartplata.mdb"
    
Set rsVrem = New ADODB.Recordset
Set rsVrem.ActiveConnection = Mconn
 
'������ ��
Combo1.Text = "�������"
Combo1.AddItem "�������"
Combo1.AddItem "���������� �������"


'��������� �������


rsVrem.Open ("SELECT Kategor.Name_Kategor FROM Kategor")
rsVrem.MoveFirst
Combo2.AddItem "���"
Do While Not rsVrem.EOF
Combo2.AddItem rsVrem.Fields("Name_Kategor")
rsVrem.MoveNext
Loop
rsVrem.Close

'�����
����� = "���"
rsVrem.Open ("SELECT KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM KLS_PODR")
rsVrem.MoveFirst
Combo3.AddItem "���"
Do While Not rsVrem.EOF
Combo3.AddItem rsVrem.Fields("NAIM_KLS")
rsVrem.MoveNext
Loop
rsVrem.Close

lblTitle.Caption = "��������� ������"
Set cnParam = Nothing
Set rsVrem = Nothing
End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub
