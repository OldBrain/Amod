VERSION 5.00
Begin VB.Form The_end 
   Caption         =   "�������� ������ "
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   Icon            =   "The_end.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4215
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin KvPay.xpcmdbutton xpcmdbutton7 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "�����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton6 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "������������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton4 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "������ �� ������ �������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "�������� ���"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton5 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      Caption         =   "������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "The_end"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True
End Sub

Private Sub xpcmdbutton1_Click()
Analizlgot.Titl = "������ ������� ������ " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 11
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, KLS_PODR.num AS ���,MainOccupant.kv_num AS ��, MainOccupant.OLDNUM AS [����], MainOccupant.BanKN AS [N ��� �� ����], MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.COMSPACE AS �������, MainOccupant.NLODGERF AS ��������� FROM MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� ORDER BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num"
'Analizlgot.�� 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(250, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton2_Click()
Analizlgot.Titl = "������ �� ������ ������� " + MainMenu.Command13.Caption + "<-> ��������� <+>-����"
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 4
Reports.sq = "SELECT Kategor.Name_Kategor AS [��������� �������], Saldo_Arh.KodKV AS ����, Saldo_Arh.SK AS ������ FROM Kategor INNER JOIN Saldo_Arh ON Kategor.��� = Saldo_Arh.KodKat"
'Analizlgot.�� 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton3_Click()
Analizlgot.Titl = "������ ������� ������� ������ " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 7
Reports.sq = "SELECT Lgota.NomNum AS ����, Lgota.Numer AS [��� ������ �� ����������� �����], Lgota.NAME_KLS AS ������������, Lgota.LPKV AS �������, Lgota.USEKV AS [������ ����������], IIf([OhteCode]=0,'���.���������������','����.�����������') AS [�������������� ������] From Lgota ORDER BY Lgota.NomNum"
'Analizlgot.�� 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 200, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 250, 200), vbBlack, True


'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton4_Click()
Analizlgot.Titl = "������ ������� " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 7
Reports.sq = "SELECT Tarif.Kategor AS [��������� �������], MainOccupant.Numer AS ����, Tarif.NameDOM AS [��� ����], Tarif.NameKV AS [��� ��������], Tarif.Value AS ����� FROM Tarif INNER JOIN MainOccupant ON (MainOccupant.DomTip = Tarif.KodDOM) AND (Tarif.KodKV = MainOccupant.KV) ORDER BY Tarif.Kategor"
'Analizlgot.�� 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 200, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 250, 200), vbBlack, True


'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show




End Sub

Private Sub xpcmdbutton5_Click()
Analizlgot.Titl = "������ ������� " + MainMenu.Command13.Caption
Analizlgot.G = 17
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS �����, KLS_PODR.Num AS ���, MainOccupant.kv_num AS ��, MainOccupant.COMSPACE AS �������, MainOccupant.NLODGERF AS ���������, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.BanKN AS [�/��], Tarif.NameDOM AS [��� ����], Tarif.NameKV AS [��� ��], Tarif.Value AS �����, Lgota.NAME_KLS AS ������, Lgota.LPKV AS �������, Lgota.USEKV AS [������ ����], IIf([OhteCode] Is Not Null,IIf([OhteCode]=0,'���.��.','����.����'),' ') AS [�������������� ������] FROM ((MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) LEFT JOIN Tarif ON (MainOccupant.DomTip = Tarif.KodDOM) AND (MainOccupant.KV = Tarif.KodKV)) LEFT JOIN Lgota ON MainOccupant.Numer = Lgota.NomNum Where (((Tarif.KodKat) = 1)) ORDER BY KLS_PODR.NAIM_KLS"
Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton6_Click()
Analizlgot.Titl = "�������� � ������������ � ������� " + MainMenu.Command13.Caption
Analizlgot.G = 10
Reports.sq = "SELECT MainOccupant.BanKN AS �����, KLS_PODR.NAIM_KLS AS �����, KLS_PODR.Num AS ���, MainOccupant.kv_num AS ��, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Lgota.NAME_KLS AS ������, MainOccupant.Priv AS ��������������� FROM (Lgota RIGHT JOIN MainOccupant ON Lgota.NomNum = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� ORDER BY KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.kv_num, MainOccupant.FAM"
Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Unload Me
Analizlgot.Show



End Sub
