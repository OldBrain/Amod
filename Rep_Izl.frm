VERSION 5.00
Begin VB.Form Rep_Izl 
   Caption         =   "����� ���������� ������"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form8"
   ScaleHeight     =   4785
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check12 
      Caption         =   "���������� ���� ��� �������"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CheckBox Check11 
      Caption         =   "���������� ������ ������� � ������ 10 �"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CheckBox Check10 
      Caption         =   "���������� ������. 10 � � ��"
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CheckBox Check9 
      Caption         =   "���������� ������� ��� ����� 10 �"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CheckBox Check8 
      Caption         =   "���������� ����������"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CheckBox Check7 
      Caption         =   "���������� ���-�� �����������"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CheckBox Check6 
      Caption         =   "���������� ����� �������"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CheckBox Check5 
      Caption         =   "���������� ��������� �������"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "���������� ����� �������"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "���������� ��������� � �������� ��"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "���������� ����� � �������� ��"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���������� ������ � �������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "��� ��������"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "������ �������� � ���������"
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "Rep_Izl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Q = "SELECT "




'SELECT KLS_PODR.NAIM_KLS, MainOccupant.FAM, Adding.Tarif AS [����� � �������� ��������], IIf(Adding!KodN=2,Adding!SummaI,0) AS [��������� �  �������� ��������], [Adding]![TarifI] AS [����� �������], IIf([Adding]![KodN]=3,[Adding]![SummaI],0) AS [��������� �������], [Adding]![ObPl] AS [����� �������], [Adding]![Propis] AS ���������, [Adding]![Socmin] AS ����������, IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0) AS [������� �*�], IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0) AS [10 � ��], IIf([������� �*�]>[10 � ��],Round([������� �*�]-[10 � ��],1),0) AS [������� ������], Round([������� ������]*([����� �������]+[����� � �������� ��������]),2) AS [��� �������]
'FROM (Adding LEFT JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���
'WHERE (((IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0))<>0) AND ((Adding.KodN)=3)) OR (((Adding.KodN)=2));
AnalizLgot.G = 2

If Check1.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "KLS_PODR.NAIM_KLS, MainOccupant.FAM" Else Q = Q + "KLS_PODR.NAIM_KLS, MainOccupant.FAM"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check2.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "Adding.Tarif AS [����� � �������� ��������]" Else Q = Q + "Adding.Tarif AS [����� � �������� ��������]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check3.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf(Adding!KodN=2,Adding!SummaI,0) AS [��������� �  �������� ��������]" Else Q = Q + "IIf(Adding!KodN=2,Adding!SummaI,0) AS [��������� �  �������� ��������]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check4.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![TarifI] AS [����� �������]" Else Q = Q + "[Adding]![TarifI] AS [����� �������]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check5.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf([Adding]![KodN]=3,[Adding]![SummaI],0) AS [��������� �������]" Else Q = Q + "IIf([Adding]![KodN]=3,[Adding]![SummaI],0) AS [��������� �������]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check6.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![ObPl] AS [����� �������]" Else Q = Q + "[Adding]![ObPl] AS [����� �������]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check7.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![Propis] AS ���������" Else Q = Q + "[Adding]![Propis] AS ���������"
AnalizLgot.G = AnalizLgot.G + 1
End If

If Check8.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![Socmin] AS ����������" Else Q = Q + "[Adding]![Socmin] AS ����������"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check9.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0) AS [������� �*�]" Else Q = Q + "IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0) AS [������� �*�]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check10.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0) AS [10 � ��]" Else Q = Q + "IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0) AS [10 � ��]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check11.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0) AS [������� ������]" Else Q = Q + "IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0) AS [������� ������]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check12.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "Round(IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0)*([Adding]![TarifI]+[Adding]![Tarif]),2) AS [��� �������]" Else Q = Q + "Round(IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0)*([Adding]![TarifI]+[Adding]![Tarif]),2) AS [��� �������]"
AnalizLgot.G = AnalizLgot.G + 1

End If


If AnalizLgot.G < 3 Then
MsgBox ("�� �� ������� ������� ������ ")
Exit Sub
End If
'If Check.Value = 1 Then If Q <> "SELECT " Then Q = Q + ", " + "" Else Q = Q + ""


Reports.sq = Q + " FROM (Adding LEFT JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���"


If Option2.Value = True Then Reports.sq = Reports.sq + " WHERE (((IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0))<>0) AND ((Adding.KodN)=3)) OR (((Adding.KodN)=2))"
If Option1.Value = True Then Reports.sq = Reports.sq + " WHERE (((IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0))<>0) AND ((Adding.KodN)=3))"

'AnalizLgot.FG1.Cols = 20
MsgBox (Reports.sq)
Unload Me
AnalizLgot.Show
AnalizLgot.�� 3

End Sub

Private Sub Form_Load()
Option1.Value = True
End Sub

Private Sub Option1_Click()
'If Option1.Value = True Then MsgBox ("true") Else MsgBox ("False")
'Check1.Value = 1

End Sub


Private Sub Option2_Click()
If Option1.Value = False Then
MsgBox ("� ������ ����������! ��� ���� ������� ����� ����� ������ �� �����, ����� ������� ����� �������� ����� � 2 ����. <<<<< ����� ���������� � ��������� ������ ���������>>>>")
'Option1.Value = True
End If
'Check1.Value = 0

End Sub
