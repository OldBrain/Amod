VERSION 5.00
Begin VB.Form RepLgota 
   BackColor       =   &H00808080&
   Caption         =   "������ �����"
   ClientHeight    =   4716
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7188
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   LinkTopic       =   "Form7"
   ScaleHeight     =   4716
   ScaleWidth      =   7188
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����������� (��� ���������)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   7215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������ � �������������, ������ ��������."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   7215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�������"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   7215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������������ ������� ����� ����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   7215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������ ������� ������� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����"
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   7215
   End
End
Attribute VB_Name = "RepLgota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEnh1_Click()
'RepParam.Show
'RepParam.Combo1.Enabled = False
'RepParam.Combo3.Enabled = False
'RepParam.Option1.Enabled = False
'RepParam.Option2.Enabled = False
'RepParam.Combo1.Text = "�����"


End Sub



Private Sub BtnEnh3_Click()
RepParam1.Show
'Analizlgot.Titl = "������" + vbNewLine + "   �� ���������� ������� � ������� �� �������-������������ ������� �������� ���������� �������" + vbNewLine + "  �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
'Analizlgot.G = 18
'Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, Adding.KodKv AS �, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Adding.ObPl AS [��� ��], Adding.Propis AS ���������, Adding.Tarif AS �����, Adding.SummaI AS ���������, Adding.SummaBl AS [��� �����], Sum([��� �����]-[���������]) AS [� ����������], tmp_lgota.NAME_KLS AS ������������, tmp_lgota.PloLG AS [��� ���], tmp_lgota.Procent AS [������� �����], [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100 AS [� ���-��], Count(tmp_lgota.UniKOd) AS [���-�� �����] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.��� = MainOccupant.Dom"

'Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.PloLG, tmp_lgota.Procent, [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)='����������') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS"


'MsgBox Reports.sq

'Analizlgot.Show

'Unload Me
'Unload RepLgota
'Unload Reports
'Exit Sub

End Sub

Private Sub Command1_Click()
RepParam1.Show
End Sub

'Private Sub Command1_Click()



'Analizlgot.Titl = Command1.Caption + " ��  " + Str(MainForm.DR)
'Analizlgot.G = 13
'Reports.sq = "AnalizLgot_k"
'Unload Me
'Analizlgot.�� 3
'Analizlgot.Show
'End Sub

Private Sub Command2_Click()
Reports.Enabled = True
Reports.Show
Unload Me
End Sub

Private Sub Command3_Click()
'MsgBox "� ���� ������ ��������� �� ��������"
'Exit Sub


Analizlgot.Titl = Command3.Caption + " ��  " + Str(MainForm.DR)
Analizlgot.G = 13
Reports.sq = "AnalizLgot_L"
Unload Me
Analizlgot.�� 3
Analizlgot.Show

End Sub

Private Sub Command4_Click()
Analizlgot.Titl = Command4.Caption + " ��  " + Str(MainForm.DR)




Analizlgot.G = 6
Reports.sq = "���������������"
Unload Me
Analizlgot.�� 0

Analizlgot.fg1.MergeCells = flexMergeRestrictAll
Analizlgot.fg1.MergeCol(-1) = True
'AnalizLgot.FG1.MergeCol(FG1.Cols - 1) = False
Analizlgot.Show

End Sub

Private Sub Command5_Click()
RepParam.Show
End Sub

Private Sub Command6_Click()
Analizlgot.Titl = "����� �� ������� �� ����� ������. ������ ������� " + Str(MainForm.DR)

Analizlgot.G = 6
Reports.sq = "�������"
Unload Me
Analizlgot.�� 3

Analizlgot.fg1.MergeCells = flexMergeRestrictAll
Analizlgot.fg1.MergeCol(-1) = True
'AnalizLgot.FG1.MergeCol(FG1.Cols - 1) = False
Analizlgot.Show

End Sub

Private Sub Command7_Click()




End Sub

Private Sub Form_Load()
Reports.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Reports.Enabled = True
End Sub
