VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Doc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8244
   ClientLeft      =   12
   ClientTop       =   252
   ClientWidth     =   12444
   ControlBox      =   0   'False
   Icon            =   "Doc.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   687
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1037
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10440
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7800
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������� �����"
      Height          =   492
      Left            =   10920
      TabIndex        =   29
      Top             =   1320
      Width           =   1332
   End
   Begin KvPay.xpcmdbutton xpcmdbutton7 
      Height          =   375
      Left            =   10200
      TabIndex        =   28
      Top             =   1680
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   656
      Caption         =   "- - -"
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
   Begin KvPay.xpcmdbutton xpcmdbutton6 
      Height          =   375
      Left            =   3120
      TabIndex        =   27
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   656
      Caption         =   "�������� Ctrl/Q"
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
   Begin KvPay.xpcmdbutton xpcmdbutton5 
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   656
      Caption         =   "������� F8"
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
   Begin KvPay.xpcmdbutton xpcmdbutton4 
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   7440
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   656
      Caption         =   "�������� F1"
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
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   480
      Width           =   2775
      _ExtentX        =   4890
      _ExtentY        =   445
      Caption         =   "�������� ��� ������ F3 "
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
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   4695
      _ExtentX        =   8276
      _ExtentY        =   445
      Caption         =   "�������� ����� F2"
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
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   480
      Width           =   1695
      _ExtentX        =   2985
      _ExtentY        =   445
      Caption         =   "������ ������ F9"
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
   Begin VB.CommandButton BtnEnh9 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�������� Ctrl/Q"
      Height          =   375
      Left            =   3120
      Picture         =   "Doc.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton BtnEnh8 
      BackColor       =   &H00BDC6BB&
      Caption         =   "������� F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton BtnEnh7 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�������� F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton BtnEnh6 
      BackColor       =   &H00BDC6BB&
      Caption         =   "- - -"
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton BtnEnh5 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton BtnEnh4 
      BackColor       =   &H00BDC6BB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      Picture         =   "Doc.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�������� Ctrl/F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�������� F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "�������� F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      Text            =   "����������"
      Top             =   840
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Doc.frx":0526
      Left            =   120
      List            =   "Doc.frx":0528
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   10095
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "1"
      Top             =   2160
      Width           =   12255
      _cx             =   21616
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   16384
      BackColorBkg    =   8421504
      BackColorAlternate=   16777215
      GridColor       =   -2147483647
      GridColorFixed  =   16711680
      TreeColor       =   8388608
      FloodColor      =   192
      SheetBorder     =   16711680
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Doc.frx":052A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483630
      ForeColorFrozen =   255
      WallPaperAlignment=   10
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
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
      Left            =   0
      TabIndex        =   21
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   12210
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   240
      Picture         =   "Doc.frx":0717
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   720
      Picture         =   "Doc.frx":0E61
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   480
      Picture         =   "Doc.frx":15AB
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   0
      Width           =   285
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
      Left            =   1080
      Picture         =   "Doc.frx":1CF5
      Top             =   240
      Width           =   156
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "���-�� ����� ���������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "� ����:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   10
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#�##0,00"
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
      Height          =   240
      Left            =   8880
      TabIndex        =   9
      Top             =   7440
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#�##0,00"
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
      Height          =   240
      Left            =   11520
      TabIndex        =   8
      Top             =   7440
      Width           =   720
   End
   Begin VB.Line Line6 
      X1              =   8
      X2              =   368
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "����������:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7200
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ������ 
         Caption         =   "������"
         Shortcut        =   {F11}
      End
      Begin VB.Menu �������� 
         Caption         =   "������ ��"
         Shortcut        =   {F9}
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   {F8}
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   {F12}
      End
      Begin VB.Menu ��������� 
         Caption         =   "�������� ����������"
         Shortcut        =   {F3}
      End
      Begin VB.Menu �������� 
         Caption         =   "�������� �����"
         Shortcut        =   {F2}
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Shortcut        =   ^Q
      End
      Begin VB.Menu �����1 
         Caption         =   "����� �� N �����"
         Shortcut        =   {F5}
      End
      Begin VB.Menu �������� 
         Caption         =   "�������� �� ���.�����"
         Shortcut        =   ^{F5}
      End
   End
End
Attribute VB_Name = "Doc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_Adding, Rs_Combo, Rs_Combo2, Cmb, CMB1 As ADODB.Recordset
Dim Rs_Combo1 As ADODB.Recordset
Dim Rs_set As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim RsRep As ADODB.Recordset
Dim Cl As String
Dim C2 As Long
Dim R As Long
Dim EditItem As String
Dim Sc As String
'Dim mconn As ADODB.Connection
Dim sq1, Kod, t As String
Dim s, Nm, Nd, SS, SumN, Kol As Double
Dim Xl As Excel.Application             'Microsoft Excel ?? Object Library
Dim bk As Excel.Workbook
Dim sh As Excel.Worksheet
Public Osnov As String
Public nu As String
Public en As Integer



Private Sub BtnEnh1__Click()

End Sub

Private Sub BtnEnh1_Click()
'FG.Enabled = False
Combo3.Enabled = True
Combo3.SetFocus
SendKeys "{F4}"
End Sub

Private Sub BtnEnh2_Click()

Combo2.Enabled = True
Combo2.SetFocus
SendKeys "{F4}"
End Sub

Private Sub BtnEnh21_Click()

End Sub

Private Sub BtnEnh3_Click()

For rw = 1 To Fg.Rows - 1
' �������� ����������
If Fg.TextMatrix(rw, 3) = "" Or Fg.TextMatrix(rw, 3) < 0 Then
MsgBox "����������� �������� ���� ����������! �������� ������"
Fg.Row = rw
Fg.Col = 3
Fg.Cell(flexcpBackColor, rw, 3, rw, 3) = vbRed
Exit Sub
End If

'�������� �������
If Fg.TextMatrix(rw, 6) = "" Or Fg.TextMatrix(rw, 6) = "........." Then
MsgBox "������������� ���.�����! �������� ������"
Fg.Row = rw
Fg.Col = 6
Fg.Cell(flexcpBackColor, rw, 6, rw, 6) = vbYellow
Exit Sub
End If

If Fg.TextMatrix(rw, 5) = "" Or Fg.TextMatrix(rw, 5) = 0 Then
MsgBox "������������� ���.�����! �������� ������"
Fg.Row = rw
Fg.Col = 5
Fg.Cell(flexcpBackColor, rw, 5, rw, 6) = vbYellow
Exit Sub
End If

Next




Doc.Enabled = False
Pod.Show
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = 500
Pod.Label1 = "�������� ������ �� ��������� � ������� �����"
Pod.Label1.Refresh
��������_Click
RS.Requery
Set Fg.DataSource = RS
����
End Sub


Private Sub BtnEnh4_Click()
PrintW.Show
    
        
     With PrintW.VP
        PrintW.VP.StartDoc
        .FontSize = 12
        .Paragraph = Text1 + " " + Label8 + " " + Label6 + " " + Label7 + " " + Label5
        .Paragraph = ""
        .FontSize = 8
        .RenderControl = Fg.hwnd
        .EndDoc
       End With
End Sub

Private Sub BtnEnh41_Click()

End Sub

Private Sub BtnEnh5_Click()
�������_Click
End Sub

Private Sub BtnEnh51_Click()

End Sub

Private Sub BtnEnh6_Click()
Text1.Enabled = True
Text1 = InputBox("������� ����� ���������� � ���������", "����������", Text1)
End Sub


Private Sub BtnEnh61_Click()

End Sub

Private Sub BtnEnh7_Click()
��������_Click
End Sub

Private Sub BtnEnh71_Click()

End Sub

Private Sub BtnEnh8_Click()
�������_Click
End Sub

Private Sub BtnEnh81_Click()

End Sub

Private Sub BtnEnh9_Click()
��������_Click
End Sub

Private Sub BtnEnh91_Click()

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Fg.SetFocus
Combo2.Enabled = False
Fg.Row = Fg.Rows - 1
Fg.Col = 3

End If
End Sub

Private Sub Combo2_LostFocus()
Fg.SetFocus
Combo2.Enabled = False
Combo3.Enabled = False
Fg.TextMatrix(Fg.Row, 3) = Str(Val(Combo2.Text))

End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
Label2 = Combo2.Text

Fg.Col = 2
Fg.Row = 1
Combo2.Enabled = False
Combo3.Enabled = False
'SendKeys "{Enter}"
End Sub

Private Sub Combo3_Click()
'Fg.TextMatrix(Fg.Row, 14) = CMB1("���")
Fg.TextMatrix(Fg.Row, 14) = Str(Val(Left(Combo3.Text, InStr(1, Combo3.Text, " ", vbTextCompare))))
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
'FG.Enabled = True
Fg.SetFocus
Combo2.Enabled = False
Fg.Row = Fg.Rows - 1
Fg.Col = 6
End If

End Sub

Private Sub Combo3_LostFocus()
Fg.SetFocus
Combo2.Enabled = False
Combo3.Enabled = False
��������
End Sub


Private Sub Combo3_Validate(Cancel As Boolean)
Label1 = Combo3.Text
Rs_Combo1.Close
��������
Fg.SetFocus
Label2 = Combo3.Text
Fg.Col = 2
Fg.Row = 1
Combo2.Enabled = False
Combo3.Enabled = False
'SendKeys "{Enter}"
End Sub


Private Sub Command1_Click()
Me.en = 0
Osnovanie.Show 1, Me
If Me.en = 10 Then Exit Sub


Set RsRep = New ADODB.Recordset

Set Rs_set = New ADODB.Recordset

RsRep.Open ("SELECT [MainOccupant]![FAM]+' '+[MainOccupant]![IM]+' '+[MainOccupant]![OT] AS FIO, KLS_PODR.NAIM_KLS as adr, MainOccupant.kv_num, Doc.KodN, Doc.NameN, Doc.Summa, Doc.Com, Doc.Key FROM (Doc INNER JOIN KLS_PODR ON Doc.Dom = KLS_PODR.���) INNER JOIN MainOccupant ON Doc.KodKv = MainOccupant.Numer WHERE (((Doc.Key)=" + Me.Fg.TextMatrix(Me.Fg.Row, 8) + "))"), Mconn

Rs_set.Open ("SELECT Settings.NamePred FROM Settings"), Mconn

'MsgBox (sumPropis(RsRep("Summa")))


Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long

nameRP = "PO"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)
est:
 
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))

DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
'���������, ���� �� ��������� ��������� ��������� ��������� Saved � ���� ��������� �� ���� ��������� - ��������� ��;
'If DocWord.Saved = False Then DocWord.Save

WordApp.Options.CheckSpellingAsYouType = False


Set TableWord = DocWord.Tables(1)
'.Add(DocWord.Range(), 10, 2)


'�������� ����� � ������ � �������
'(�����_������, �����_�������)


'�������� �����������
'TableWord.Cell(7, 1).Range.Text = MainForm.Label3.Caption
'TableWord.Cell(3, 6).Range.Text = MainForm.Label3.Caption
TableWord.Cell(7, 1).Range.Text = Rs_set("NamePred")
TableWord.Cell(3, 6).Range.Text = Rs_set("NamePred")
Rs_set.Close
'�����
TableWord.Cell(20, 6).Range.Text = RsRep("Summa")
TableWord.Cell(20, 14).Range.Text = Int(RsRep("Summa"))
TableWord.Cell(20, 16).Range.Text = Right(Round(RsRep("Summa") - Int(RsRep("Summa")), 2), 2)

' ������� ��
TableWord.Cell(12, 9).Range.Text = RsRep("Fio") + " �����:" + RsRep("Adr") + " ��.�" + RsRep("kv_num")
TableWord.Cell(22, 2).Range.Text = RsRep("Fio") + " �����:" + RsRep("Adr") + " ��.�" + RsRep("kv_num")
'����� ��������
TableWord.Cell(26, 2).Range.Text = sumPropis(RsRep("Summa"))
TableWord.Cell(22, 7).Range.Text = sumPropis(RsRep("Summa"))
'����
TableWord.Cell(13, 3).Range.Text = Date
TableWord.Cell(10, 8).Range.Text = Day(Date)
TableWord.Cell(10, 10).Range.Text = Choose(Month(Date), "������", "�������", "�����", "������", "���", "����", "����", "�������", "��������", "�������", "������", "�������")
TableWord.Cell(10, 12).Range.Text = Year(Date)
TableWord.Cell(28, 10).Range.Text = Day(Date)
TableWord.Cell(28, 12).Range.Text = Choose(Month(Date), "������", "�������", "�����", "������", "���", "����", "����", "�������", "��������", "�������", "������", "�������")
TableWord.Cell(28, 14).Range.Text = Year(Date)

'�����

TableWord.Cell(13, 2).Range.Text = nu
TableWord.Cell(9, 8).Range.Text = nu
'���������
TableWord.Cell(24, 2).Range.Text = Me.Osnov
TableWord.Cell(14, 7).Range.Text = Me.Osnov


Set DocWord = Nothing

'���������� ������ - Word
Set WordApp = Nothing


End Sub

Private Sub Command2_Click()
Dim lineFile As String
Dim RsData As ADODB.Recordset
Dim Da As String



Nd = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 1)

FileName = App.Path + "/fg/reestr" + Nd + ".csv"

Set RsData = New ADODB.Recordset
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile(FileName, True)

RsData.Open "SELECT doc.DataR, doc.NameN, doc.Summa, MainOccupant.TELEPHONE, doc.Cod, LTrim([MainOccupant]![FAM])+' '+Left([MainOccupant]![IM],1)+'. '+Left([MainOccupant]![OT],1)+'. '+[NAIM_KLS]+' �� '+[MainOccupant]![kv_num] AS FIO FROM (doc INNER JOIN MainOccupant ON doc.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((doc.Cod)=" + Nd + "))", Mconn

'������ ������ �����
lineFile = "����;�����;������������ ������;�������;���/�����"
a.WriteLine (lineFile)



RsData.MoveFirst

Do While Not RsData.EOF

'1. ����
Da = CStr(RsData("DataR"))
'2. �����
su = Str(RsData("Summa"))
'3. ������������ ������
NameUs = RsData("NameN")
'4.�������
Phone = RsData("TELEPHONE")
'5. �������/�����
FIO = RsData("FIO")



lineFile = Da + ";" + su + ";" + RsData("NameN") + ";" + RsData("TELEPHONE") + ";" + FIO
'Format(RsData("Sum-SummaI"), "###0.00")

a.WriteLine (lineFile)

RsData.MoveNext
Loop
a.Close

End Sub

Private Sub FG_AfterDataRefresh()
Fg.ColComboList(11) = "..."
Fg.ColComboList(13) = "..."
�������
End Sub

Private Sub FG_Click()
Combo3.Enabled = False
Combo2.Enabled = False
Text1.Enabled = False

CMB1.MoveFirst
Do While Not CMB1.EOF
If CMB1("���") = Fg.TextMatrix(Fg.Row, 14) Then
Combo3.Text = CStr(CMB1("���")) & "  " & CMB1("Naim_kls") & " ��� � " & CMB1("Num")
Exit Do
End If
CMB1.MoveNext
Loop

'Cl = CStr(CMB1("���")) & "  " & CMB1("Naim_kls") & " ��� � " & CMB1("Num")
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, Shift As Integer)
CMB1.MoveFirst
Do While Not CMB1.EOF
If CMB1("���") = Fg.TextMatrix(Fg.Row, 14) Then
Combo3.Text = CStr(CMB1("���")) & "  " & CMB1("Naim_kls") & " ��� � " & CMB1("Num")
Exit Do
End If
CMB1.MoveNext
Loop
End Sub

Private Sub FG_KeyPress(KeyAscii As Integer)
'MsgBox (Str(KeyAscii))
If KeyAscii = 27 Then �������_Click

If KeyAscii = 32 Then
If Fg.TextMatrix(Fg.Row, 9) = 0 Then
Fg.TextMatrix(Fg.Row, 9) = 1
Fg.Cell(flexcpForeColor, Fg.Row, 1, Fg.Row, 10) = vbBlue
Fg.Cell(flexcpFontBold, Fg.Row, 1, Fg.Row, 10) = True
Else
Fg.TextMatrix(Fg.Row, 9) = 0
Fg.Cell(flexcpForeColor, Fg.Row, 1, Fg.Row, 10) = vbBlack
Fg.Cell(flexcpFontBold, Fg.Row, 1, Fg.Row, 10) = False
End If
End If



End Sub


Private Sub FG_KeyUp(KeyCode As Integer, Shift As Integer)
CMB1.MoveFirst
Do While Not CMB1.EOF
If CMB1("���") = Fg.TextMatrix(Fg.Row, 14) Then
Combo3.Text = CStr(CMB1("���")) & "  " & CMB1("Naim_kls") & " ��� � " & CMB1("Num")
Exit Do
End If
CMB1.MoveNext
Loop
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Arhiv = True And Fg.TextMatrix(0, Fg.Col) <> "...." Then Cancel = True
R = Row
C2 = Col
EditItem = Fg.TextMatrix(Row, Col)
End Sub

Private Sub lblTitle_Click()
'FSize Me
ChangeState Me
End Sub

Private Sub Text1_DblClick()
Text1.Enabled = True
Text1 = InputBox("", "����������", Text1)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)

ReestrDoc.Fg.TextMatrix(ReestrDoc.Fg.Row, 4) = Text1
End Sub

' ��������� ����
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    'On Error Resume Next
    Select Case Button.KEY
        Case "New"
            ��������_Click
        Case "Delete"
            �������_Click
        Case "Save"
            �������_Click
    End Select
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)



If Fg.TextMatrix(Fg.Row, 7) = "" Then Fg.TextMatrix(Fg.Row, 7) = 0

If SumN <> Fg.TextMatrix(Fg.Row, 7) And Fg.TextMatrix(Fg.Row, 10) = 1 Then
Mconn.Execute ("UPDATE Adding INNER JOIN Doc ON Adding.KodDoc = Doc.Key SET Adding.SummaI = [Doc]![Summa] WHERE (((Adding.KodDoc)=" + Fg.TextMatrix(Fg.Row, 8) + "))")
Else
End If

If Fg.TextMatrix(Fg.Row, 7) = "" Then Fg.TextMatrix(Fg.Row, 7) = 0

If Fg.TextMatrix(Fg.Row, Fg.Col) = "" Then Exit Sub

Rs_Combo.MoveFirst
Do While Not Rs_Combo.EOF
         If Rs_Combo("Kod") = Fg.TextMatrix(Fg.Row, 3) Then
    Fg.TextMatrix(Fg.Row, 4) = Rs_Combo("Naim")
    Fg.TextMatrix(Fg.Row, 12) = Rs_Combo("Tip")
                  End If
Rs_Combo.MoveNext
Loop

If Fg.TextMatrix(Fg.Row, 5) = "" Then Fg.TextMatrix(Fg.Row, 5) = 0
Q = "SELECT MainOccupant.Numer, MainOccupant.FAM,MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.DOM From MainOccupant WHERE(((MainOccupant.Numer)=" + Fg.TextMatrix(Fg.Row, 5) + "))"
Rs_Combo2.Open (Q)
Fg.TextMatrix(Fg.Row, 14) = Str(Val(Left(Combo3.Text, InStr(1, Combo3.Text, " ", vbTextCompare))))

Fg.ComboList = ""
Rs_Combo2.Close

'---------------------------------------------------
On Error GoTo �����
Rs_Combo1.MoveFirst
Do While Not Rs_Combo1.EOF
'MsgBox (FG.TextMatrix(FG.Row, 6) + "  " + Rs_Combo1.Fields("FAM").Value + " " + Rs_Combo1.Fields("IM").Value + " " + Rs_Combo1.Fields("OT").Value + "�� � " + Rs_Combo1.Fields("kv_num").Value)
If Fg.TextMatrix(Fg.Row, 6) = Rs_Combo1.Fields("FAM").Value + " " + Rs_Combo1.Fields("IM").Value + " " + Rs_Combo1.Fields("OT").Value + "�� � " + Rs_Combo1.Fields("kv_num").Value Then
Fg.TextMatrix(Fg.Row, 5) = Rs_Combo1.Fields("Numer").Value
End If
Rs_Combo1.MoveNext
Loop

'����
If Kol < 1 Then
Mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + Fg.TextMatrix(1, 1) + "))")
Kol = Kol + 1
End If
Exit Sub

�����:
If Err.Number = 3021 Then
MsgBox ("�� ����� ������ ��� �������!")
Combo3.Enabled = True
Combo3.SetFocus
Else
MsgBox ("������: " + Err.Description)
End If

End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


If Col = 1 Or Col = 4 Or Col = 5 Then Cancel = True
If Col = 2 Then Fg.ComboList = ""
If Col = 7 Then Fg.ComboList = ""

If Fg.TextMatrix(Fg.Row, 7) <> "" Then SumN = Val(Fg.TextMatrix(Fg.Row, 7))
If Fg.TextMatrix(Fg.Row, 7) <> "" Then SS = Fg.TextMatrix(Fg.Row, 7) Else SS = 0

'If FG.TextMatrix(Row, 7) <> 0 Then
If Fg.TextMatrix(Row, 10) <> 0 Then
If Col <> 7 And Col <> 13 And Col <> 11 Then
Cancel = True
End If
End If

' ����������
If Col = 3 Then
Cl = ""
If Combo2.Text = "����� ����������" Then
Rs_Combo.MoveFirst
Do While Not Rs_Combo.EOF
Cl = Cl + CStr(Rs_Combo("Kod")) & vbTab & Rs_Combo("Naim") + "|"
Rs_Combo.MoveNext
Loop
Fg.ComboList = Cl
Else
Fg.ComboList = ""

                             End If

 End If
 
'�������

CMB1.MoveFirst
Do While Not CMB1.EOF
If CMB1("���") = Fg.TextMatrix(Fg.Row, 14) Then
Combo3.Text = CStr(CMB1("���")) & "  " & CMB1("Naim_kls") & " ��� � " & CMB1("Num")
Exit Do
End If
CMB1.MoveNext
Loop


��������




If Fg.TextMatrix(0, Fg.Col) = "�.�.�." Then
Cl = ""
On Error GoTo �����
Rs_Combo1.MoveFirst
J = 0
Do While Not Rs_Combo1.EOF
J = J + 1
If J > 1000 Then
Msg.Show
Msg.Label1.Caption = "��������� ���������� ����� � ����� ���������"


BtnEnh1_Click

'MainMenu.Enabled = True
'Unload Me
Exit Sub
End If
If Rs_Combo1("���") <> "" Then Cl = Cl + CStr(Rs_Combo1("���")) + "�� � " + Rs_Combo1.Fields("kv_num").Value & vbTab & CStr(Rs_Combo1("Numer")) + "|"
Rs_Combo1.MoveNext
Loop
Fg.ComboList = Cl

End If


If Col = 3 Then

If Fg.TextMatrix(Row, 10) <> 0 Then Cancel = True
End If

If Col = 5 Or Col = 6 Then
If Fg.TextMatrix(Row, 10) <> 0 Then Cancel = True
End If


����
Exit Sub

�����:
If Err.Number = 3021 Then
MsgBox ("�� ����� ������ ��� �������!")
Combo3.Enabled = True
Combo3.SetFocus
Else
MsgBox ("������: " + Err.Description)
End If


End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim nr As Long, nc As Long      '��� ������ �������� ���� ����������� � ������ � �������
    Static R As Long, c As Long     '��� �� ���������� ��� �������� ������� ������
    nr = Fg.MouseRow:    nc = Fg.MouseCol  ' get coordinates
    
    If nr < 1 Or nc = -1 Then
    Fg.ToolTipText = ""
    Exit Sub
    End If
    If c <> nc Or R <> nr Then                   ' update tooltip text
        
       If Fg.TextMatrix(nr, nc) <> "" Then
        Fg.ToolTipText = Fg.TextMatrix(nr, nc)
        End If
        R = nr:            c = nc
        DoEvents
    End If
End Sub

    
'End Sub


Private Sub Form_Load()

MakeWindow Me, True



If Arhiv = True Then BtnEnh2.Enabled = False
If Arhiv = True Then BtnEnh1.Enabled = False
If Arhiv = True Then BtnEnh3.Enabled = False
If Arhiv = True Then BtnEnh6.Enabled = False
If Arhiv = True Then BtnEnh7.Enabled = False
If Arhiv = True Then BtnEnh8.Enabled = False
If Arhiv = True Then BtnEnh9.Enabled = False


Kol = 0
'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.amd;Persist Security Info=True"
'mconn.Open "data/kvartplata.amd"
'Set rs_Tit = New ADODB.Recordset
Combo3.Enabled = False
Combo2.Enabled = False
Text1.Enabled = False


'Recordset ��� ��������
Set Cmb = New ADODB.Recordset
Set Cmb.ActiveConnection = Mconn
Set CMB1 = New ADODB.Recordset
Set CMB1.ActiveConnection = Mconn
Cmb.CursorType = adOpenForwardOnly
Cmb.LockType = adLockBatchOptimistic
Cmb.Open "Nachisleniy"
CMB1.Open "SELECT KLS_PODR.* FROM KLS_PODR ORDER BY KLS_PODR.NAIM_KLS"

'FG.BackColorFrozen = RGB(200, 255, 200)

' ��������� Combo2 ��� ����������
'Set Combo2.DataSource = Combo_RS
Combo2.Text = ReestrDoc.Fg.TextMatrix(ReestrDoc.Fg.Row, 3)
Cl = "����� ����������"
Cmb.MoveFirst
Do While Not Cmb.EOF
Combo2.AddItem Cl
Cl = CStr(Cmb("Kod")) & "  " & Cmb("Naim")
'codN(Combo_RS("Kod")) = Combo_RS("Kod")
Cmb.MoveNext
Loop

' ��������� Combo3 ��� �������
'Set Combo2.DataSource = Cmb1

Combo3.Text = ReestrDoc.Fg.TextMatrix(ReestrDoc.Fg.Row, 10)
'cl = "0   ��� ����  0"
CMB1.MoveFirst
Do While Not CMB1.EOF
If CMB1("���") <> 0 Then
Cl = CStr(CMB1("���")) & "  " & CMB1("Naim_kls") & " ��� � " & CMB1("Num")
Combo3.AddItem Cl
End If
CMB1.MoveNext
Loop

Set RS = New ADODB.Recordset
Set RS.ActiveConnection = Mconn

Set Rs_Combo = New ADODB.Recordset
Set Rs_Combo.ActiveConnection = Mconn

Set Rs_Combo1 = New ADODB.Recordset
Set Rs_Combo1.ActiveConnection = Mconn

Set Rs_Combo2 = New ADODB.Recordset
Set Rs_Combo2.ActiveConnection = Mconn

'Doc.Caption
lblTitle.Caption = "�������� �� ����������/���������/�������� �� ���� " + ReestrDoc.Fg.TextMatrix(ReestrDoc.Fg.Row, 2)

Fg.Editable = flexEDKbdMouse

Label1 = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 10)
Label2 = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 3)
Text1 = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 4)
'rs_Tit.Open

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockBatchOptimistic

Rs_Combo.CursorType = adOpenForwardOnly
Rs_Combo.LockType = adLockBatchOptimistic

Rs_Combo2.CursorType = adOpenForwardOnly
Rs_Combo2.LockType = adLockBatchOptimistic

Kod = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 1)

RS.Open ("SELECT Doc.*, Doc.Cod From Doc WHERE (((Doc.Cod)=" + Kod + "))")
Rs_Combo.Open "Nachisleniy  ORDER BY nachisleniy.Kod DESC"




 '��� ����� Recordset ��� Combo �������, ������������ �� ����������
 '������ � ����� ���������
 
  
sq1 = "SELECT MainOccupant.Numer,MainOccupant.FAM,MainOccupant.IM,MainOccupant.OT, MainOccupant.kv_num, MainOccupant!FAM+" & Chr(34) & " " & Chr(34) + "+MainOccupant!IM+" + Chr(34) + " " + Chr(34) + " + MainOccupant!OT " + " AS ���, "
'MsgBox (sq1)
sq1 = sq1 & Chr(34) & "��." & Chr(34) & "+MainOccupant.Kv_Num+" & Chr(34)
sq1 = sq1 + "��� � " & Chr(34) & "+KLS_PODR!Num AS ���, MainOccupant.Dom FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom=KLS_PODR.���"
Kod1 = ""
'If Val(ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)) <> 0 Then
If Val(Combo3.Text) <> 0 Then
'Kod = Str(Val(ReestrDoc.FG.TextMatrix(ReestrDoc.r, 10)))
'Kod1 = Str(Val(Combo3.Text))
Kod1 = Str(Val(Left(Combo3.Text, InStr(1, Combo3.Text, " ", vbTextCompare))))
sq1 = sq1 + " WHERE (((MainOccupant.Dom)=" + Kod1 + ")) ORDER BY MainOccupant.FAM"
End If

Rs_Combo1.Open (sq1)









Fg.DataMode = flexDMBoundImmediate
' C�������, �������� ����������� ��� ����������
'    FG.AllowUserResizing = flexResizeBoth
 '   FG.ExtendLastCol = True
    Fg.ExplorerBar = flexExSort
    Fg.AutoSearch = flexSearchFromCursor


Set Fg.DataSource = RS
����

'�����������
'FG.MergeCells = flexMergeRestrictAll
'FG.MergeCol(-1) = True
'FG.MergeCol(FG.Cols - 1) = False




Fg.ColComboList(11) = "..."
Fg.ColComboList(13) = "..."
����
'ChangeState Me
'FSize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'ReestrDoc.Show
'ReestrDoc.FG.Refresh
'ReestrDoc.FG.Redraw
'Load ReestrDoc
End Sub

Private Sub xpcmdbutton1_Click()
��������_Click
End Sub

Private Sub xpcmdbutton2_Click()
Combo3.Enabled = True
Combo3.SetFocus
SendKeys "{F4}"
End Sub

Private Sub xpcmdbutton3_Click()
Combo2.Enabled = True
Combo2.SetFocus
SendKeys "{F4}"
End Sub

Private Sub xpcmdbutton4_Click()
��������_Click
End Sub

Private Sub xpcmdbutton5_Click()
�������_Click

End Sub

Private Sub xpcmdbutton6_Click()
��������_Click
End Sub

Private Sub xpcmdbutton7_Click()
Text1.Enabled = True
Text1 = InputBox("������� ����� ���������� � ���������", "����������", Text1)
End Sub

Private Sub ��������_Click()

RS.AddNew
RS("doc.Cod") = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 1)
RS("DataR") = ReestrDoc.Fg.TextMatrix(ReestrDoc.R, 2)
RS("NameKv") = "........."
RS("DOM") = Str(Val(Left(Combo3.Text, InStr(1, Combo3.Text, " ", vbTextCompare))))



If Combo2.Text <> "����� ����������" Then
RS("KodN") = Val(Combo2.Text)

'ReestrDoc.FG.TextMatrix(ReestrDoc.R, 8)
Else
RS("KodN") = -1
End If

RS.UpdateBatch
Fg.DataRefresh
RS.MoveLast

Rs_Combo.MoveFirst
Do While Not Rs_Combo.EOF
If Rs_Combo("Kod") = Fg.TextMatrix(Fg.Row, 3) Then Fg.TextMatrix(Fg.Row, 4) = Rs_Combo("Naim")
Rs_Combo.MoveNext
Loop
If Fg.TextMatrix(Fg.Row, 5) = "" Then Fg.TextMatrix(Fg.Row, 5) = 0
Mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + Fg.TextMatrix(1, 1) + "))")


End Sub

Private Sub �������_Click()

For rw = 1 To Fg.Rows - 1
' �������� ����������
If Fg.TextMatrix(rw, 3) = "" Or Fg.TextMatrix(rw, 3) < 0 Then
MsgBox "����������� �������� ���� ����������! �������� ������"
Fg.Row = rw
Fg.Col = 3
Fg.Cell(flexcpBackColor, rw, 3, rw, 3) = vbRed
Exit Sub
End If

'�������� �������
If Fg.TextMatrix(rw, 6) = "" Or Fg.TextMatrix(rw, 6) = "........." Then
MsgBox "������������� ���.�����! �������� ������"
Fg.Row = rw
Fg.Col = 6
Fg.Cell(flexcpBackColor, rw, 6, rw, 6) = vbYellow
Exit Sub
End If

If Fg.TextMatrix(rw, 5) = "" Or Fg.TextMatrix(rw, 5) = 0 Then
MsgBox "������������� ���.�����! �������� ������"
Fg.Row = rw
Fg.Col = 5
Fg.Cell(flexcpBackColor, rw, 5, rw, 6) = vbYellow
Exit Sub
End If


Next


If Fg.Rows > 1 Then Kod1 = Fg.TextMatrix(1, 1)
st = Doc.Label5
ad = Chr(34) + Combo3.Text + Chr(34)

'MsgBox (Label5)
'Unload ReestrDoc
'st = Str(Int(s)) + "," + Str(s - Int(s))
'On Error Resume Next



If Kod1 <> "" Then
Mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Summa = " + st + ",ReestrDoc.Adres = " + ad + ",ReestrDoc.coment = " + Chr(34) + Text1 + Chr(34) + "  WHERE (((ReestrDoc!Cod)=" + Kod1 + "))")
End If
'ReestrDoc.FG.TextMatrix(ReestrDoc.r, 5) = s
'ReestrDoc.FG.TextMatrix(ReestrDoc.r, 4) = Doc.Label5
ReestrDoc.Hide
Unload Doc
Unload ReestrDoc
Load ReestrDoc
ReestrDoc.Show
ReestrDoc.Fg.DataRefresh
ReestrDoc.Refresh


End Sub

Private Sub ��������_Click()
BtnEnh1_Click
End Sub

Private Sub ���������_Click()
BtnEnh2_Click
End Sub

Private Sub ��������_Click()
frmSelectMonth.txtYear = Year(Fg.TextMatrix(Fg.Row, 15))
frmSelectMonth.Show 1




'MsgBox Fg.TextMatrix(Fg.Row, 8)
'MsgBox Fg.TextMatrix(Fg.Row, 15)
'MsgBox Month(Doc.Fg.TextMatrix(Doc.Fg.Row, 15))
End Sub

Private Sub ��������_Click()
Fg.TextMatrix(R, C2) = EditItem

Fg.Cell(flexcpForeColor, R, C2, R, C2) = vbRed
Fg.Cell(flexcpFontBold, R, C2, R, C2) = True

End Sub

Private Sub ������_Click()

Nd = Fg.TextMatrix(1, 1)
'sq = ""
Analizlgot.G = 11




Analizlgot.Titl = "�������� ������ �" + Nd + "/" + Text1 + " /" + Label8 + " " + Label6 + " /" + Label7 + " " + Label5



Analizlgot.StrSQL = "SELECT KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, Doc.NameN AS ���, Doc.Summa AS ��������, Doc.DataR AS [���� ������], Doc.RealData AS [������ ������], Doc.Com AS ���������� FROM (Doc INNER JOIN MainOccupant ON Doc.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((Doc.Cod)=" + Nd + "))"
Analizlgot.Show








'Analizlgot.FG1.Subtotal flexSTSum, 4, 10, , RGB(250, 250, 200), vbBlack, True, "� ���� �/��:"
'Analizlgot.FG1.Subtotal flexSTSum, 4, 11, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 12, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 13, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 4, 14, , RGB(250, 250, 200), vbBlack, True





Unload Me
'Analizlgot.�� 1

End Sub

Private Sub �����1_Click()
Sc = InputBox("������� ����� �������� �����")

End Sub

Private Sub ��������_Click()
Dim RazDoc_Adding As ADODB.Recordset
If Fg.Row <> 0 Then
Nd = Fg.TextMatrix(1, 1)
Else
Pod.Label1 = "��� ������ ��� ��������"
Pod.Command1.Visible = True
Exit Sub
End If
Doc.Enabled = False
Pod.Show

'����������� ������ �� ������ ���� � ���� ���� ����������� ����
'Mconn.Execute ("UPDATE Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV) SET Adding.SaldoN = [Saldo_Arh]![SK] WHERE ((([Saldo_Arh]![SK]-[Adding]![SaldoN])<>0))")
Mconn.Execute ("UPDATE Adding LEFT JOIN Saldo_Arh ON (Adding.KodKv = Saldo_Arh.KodKV) AND (Adding.KodKat = Saldo_Arh.KodKat) SET Adding.SaldoN = [Saldo_Arh]![SK] WHERE ((([Saldo_Arh]![SK]-[Adding]![SaldoN])<>0))")

Set RazDoc_Adding = New ADODB.Recordset
Set RazDoc_Adding.ActiveConnection = Mconn
RazDoc_Adding.CursorType = adOpenStatic
'adOpenForwardOnly
RazDoc_Adding.LockType = adLockBatchOptimistic
'������� ��������� ������
Mconn.Execute ("DELETE Doc.Cod, Doc.KodN From Doc WHERE (((Doc.Cod)=" + Nd + ") AND ((Doc.KodN)=-1)) OR (((Doc.Cod)=" + Nd + ") AND ((Doc.KodN) Is Null))")

'��������� ������ ����������
Mconn.Execute ("UPDATE Doc SET Doc.Com =" + Chr(34) + " " + Chr(34) + "WHERE (((Doc.Com) Is Null)) And (((Doc.Cod)=" + Nd + "))")

Pod.ProgressBar1.Value = 50
'RazDoc_Adding.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.KeyAdding, Doc.Stst, Doc.Com, Doc.Tip, Doc.Button FROM Doc LEFT JOIN Adding ON Doc.Key = Adding.KodDoc WHERE (((Doc.Cod)=" + Nd + ") AND ((Adding.KodDoc) Is Null))")
'********************************
'If RazDoc_Adding("Cod") <> "" Then RazDoc_Adding.MoveFirst
'Do While Not RazDoc_Adding.EOF

'Mconn.Execute ("INSERT INTO Adding ( DataR, KodN, NameN, KodKv, SummaI, KodDoc, ispr, Com, Tip ) SELECT RazDoc_Adding.DataR, RazDoc_Adding.KodN, RazDoc_Adding.NameN, RazDoc_Adding.KodKv, RazDoc_Adding.Summa, RazDoc_Adding.Key, RazDoc_Adding.Stst, RazDoc_Adding.Com, RazDoc_Adding.Tip FROM RazDoc_Adding")
'Mconn.Execute ("INSERT INTO Adding ( DataR, KodN, NameN, KodKv, SummaI, KodDoc, ispr, Com, Tip ) SELECT RazDoc_Adding.DataR, RazDoc_Adding.KodN, RazDoc_Adding.NameN, RazDoc_Adding.KodKv, RazDoc_Adding.Summa, RazDoc_Adding.Key, RazDoc_Adding.Stst, RazDoc_Adding.Com, RazDoc_Adding.Tip From RazDoc_Adding WHERE (((RazDoc_Adding.Cod)=" + Nd + "))")

Mconn.Execute ("INSERT INTO Adding ( DataR, KodN, NameN, KodKv, SummaI, KodDoc, ispr, Com, Tip, DataT ) SELECT RazDoc_Adding.DataR, RazDoc_Adding.KodN, RazDoc_Adding.NameN, RazDoc_Adding.KodKv, RazDoc_Adding.Summa, RazDoc_Adding.Key, RazDoc_Adding.Stst, RazDoc_Adding.Com, RazDoc_Adding.Tip, RazDoc_Adding.Realdata From RazDoc_Adding WHERE (((RazDoc_Adding.Cod)=" + Nd + "))")


'RazDoc_Adding.MoveNext
'Loop
'RazDoc_Adding.Close
Pod.ProgressBar1.Value = 100
���������� = Chr(34) + " ��� �" + Str(Nd) + Chr(34)
Mconn.Execute ("UPDATE Adding INNER JOIN Doc ON Adding.KodDoc = Doc.Key SET Adding.KodKv = [Doc]![KodKv], Adding.SummaI = [Doc]![Summa], Adding.KodN = [Doc]![KodN], Adding.NameN = [Doc]![NameN], Adding.Com = " + ���������� + " +[Doc]![Com] " + ", Adding.Tip = [Doc]![Tip], Adding.DataR = [Doc]![DataR] WHERE (((Doc.Cod)=" + Nd + "))")
Mconn.Execute ("UPDATE Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod SET Adding.KodKat = [nachisleniy]![���Kategor], Adding.NameKat = [nachisleniy]![Kategor] WHERE (((Adding.KodKat)=0) AND ((Adding.KodDoc)<>0)) OR (((Adding.KodKat) Is Null))")

'����������� ������
Pod.ProgressBar1.Value = 150


Mconn.Execute ("UPDATE (Saldo_Arh INNER JOIN Adding ON (Saldo_Arh.KodKV = Adding.KodKv) AND (Saldo_Arh.KodKat = Adding.KodKat)) LEFT JOIN Doc ON Adding.KodDoc = Doc.Key SET Adding.SaldoN = [Saldo_Arh]![SK] WHERE (((Doc.Cod)=" + Nd + "))")

'���� ������������� ������ �� ������ �� ������� ������
'������� ���������� ������ � �������� 4 ����������� ������

'mconn.Execute ("DELETE TMP_DOC.* FROM TMP_DOC")
'mconn.Execute ("INSERT INTO TMP_DOC ( KodKv, Cod, [Key], ���Kategor ) SELECT Doc.KodKv, Doc.Cod, Doc.Key, nachisleniy.���Kategor FROM Doc INNER JOIN nachisleniy ON Doc.KodN = nachisleniy.Kod WHERE (((Doc.Cod)=" + Nd + "))")
'mconn.Execute ("UPDATE Adding INNER JOIN TMP_DOC ON (Adding.KodKat = TMP_DOC.���Kategor) AND (Adding.KodKv = TMP_DOC.KodKv) SET TMP_DOC.Saldo = round([Adding]![SaldoN],2) WHERE (((Adding.KodDoc)=0))")
'mconn.Execute ("UPDATE Adding INNER JOIN TMP_DOC ON Adding.KodDoc = TMP_DOC.Key SET Adding.SaldoN = round([TMP_DOC]![Saldo],2)")


Pod.ProgressBar1.Value = 250
'��������� �������
Mconn.Execute ("UPDATE Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod SET Adding.Formula = [nachisleniy]![Formula], Adding.FormulaB = [nachisleniy]![FormulaB] WHERE (((Adding.Formula) Is Null))")
Pod.ProgressBar1.Value = 300
'mconn.Execute ("UPDATE Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod SET Adding.FormulaB = [nachisleniy]![FormulaB] WHERE (((Adding.FormulaB) Is Null))")
Pod.ProgressBar1.Value = 350
Pod.Label1.Caption = "���������� ������� �����"
Pod.Label1.Refresh
Pod.Label1.FontSize = 8
For rw = 1 To Fg.Rows - 1
Pod.Label1.Caption = "���������� ������ �/����� >" + Fg.TextMatrix(rw, 6)
Pod.Label1.Refresh

Pod.ProgressBar1.Value = 450
MainForm.���������������� Str(Fg.TextMatrix(rw, 5))

Pod.ProgressBar1.Value = 400
MainForm.RSaldoK Str(Fg.TextMatrix(rw, 5))
'Pod.Label1.FontItalic = True
'Pod.ProgressBar1.Value = 450
'MainForm.���������������� Str(FG.TextMatrix(Rw, 5))
'MainForm.RSaldoK Str(FG.TextMatrix(Rw, 5))
Next
Pod.ProgressBar1.Value = 500
Pod.Label1.FontSize = 10

'���������� ������
 '+ Nd +
Mconn.Execute ("UPDATE Doc SET Doc.Stst = 1 WHERE (((Doc.Cod)=" + Nd + "))")

If Mconn.Errors.Count = 0 Then
Pod.Label1.Caption = "������ ��������� �������."
Else
Pod.Label1.Caption = "������ ��� ���������� ������ �� ������� ������, ���������� ��������� ��������"
End If
Pod.Command1.Visible = True

Mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 1 WHERE (((ReestrDoc.Cod)=" + Nd + "))")



End Sub



Private Sub ��������1_Click()

Doc.Enabled = False
Pod.Show

Set rs_Adding = New ADODB.Recordset
Set rs_Adding.ActiveConnection = Mconn
rs_Adding.CursorType = adOpenForwardOnly
rs_Adding.LockType = adLockBatchOptimistic

'rs_Adding.Open ("SELECT Adding.KodKv FROM Adding INNER JOIN Doc ON Adding.KodDoc = Doc.Key GROUP BY Adding.KodKv")





' ������� ������

RS.MoveFirst
                                Do While Not RS.EOF
n = RS.Fields("Key").Value

Mconn.Execute ("DELETE Adding.*, Adding.KodDoc From Adding WHERE (((Adding.KodDoc)=" + Str(n) + "))")
RS.MoveNext
                                       Loop
                                       
Nd = Fg.TextMatrix(1, 1)

'MsgBox (Nd)
'Qdoc = "INSERT INTO Adding ( NameN, KodKat, KodN, KodKv, KodDoc, NameKat, DataR, Socmin, Propis, Projiv, ProLift, ObPl, PolPl, Formula, Tarif, Com, FLOOR, SchetZ, TarifD, TarifI, ispr, TipDomKod, TipKvKod, Tip, SummaI, Parametr, Lig, LgotaVid ) SELECT nachisleniy.Naim, nachisleniy.���Kategor, Doc.KodN, Doc.KodKv, Doc.Key, nachisleniy.Kategor, Doc.DataR, Socmin.Value, MainOccupant.NLODGERF, MainOccupant.NLODGER, MainOccupant.NLODLIFT, MainOccupant.COMSPACE, MainOccupant.HABSPACE, nachisleniy.Formula, Tarif.Value, Doc.Com, MainOccupant.FLOOR, nachisleniy.SchetZ, Tarif.TarifD, Tarif.TarifI, Doc.Stst, Tarif.KodDOM, Tarif.KodKV, nachisleniy.Tip, Doc.Summa, " + Chr(34) + "�� ����������" + Chr(34) + " AS ���������1, nachisleniy.Lig, nachisleniy.Vid "

'mconn.Execute (Qdoc + "FROM (Socmin INNER JOIN (MainOccupant INNER JOIN (nachisleniy INNER JOIN Doc ON nachisleniy.Kod = Doc.KodN) ON MainOccupant.Numer = Doc.KodKv) ON (Socmin.koli = MainOccupant.NLODGERF) AND (Socmin.Kategor = nachisleniy.Kategor)) INNER JOIN Tarif ON (MainOccupant.KV = Tarif.KodKV) AND (nachisleniy.���Kategor = Tarif.KodKat) AND (MainOccupant.DomTip = Tarif.KodDOM) WHERE (((Doc.Cod)=" + Nd + "))")
                                       
 Mconn.Execute ("INSERT INTO Adding ( DataR, KodN, NameN, KodKv, SummaI, KodDoc, Tip, Com, ispr ) SELECT Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.Summa, Doc.Key, Doc.Tip, Doc.Com, Doc.Stst From Doc WHERE (((Doc.Cod)=" + Nd + "))")
 
 RS.MoveFirst
                                Do While Not RS.EOF
n = RS.Fields("Key").Value

'mconn.Execute ("DELETE Adding.*, Adding.KodDoc From Adding WHERE (((Adding.KodDoc)=" + Str(N) + "))")
Mconn.Execute ("UPDATE (Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod) INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.KodKat = [nachisleniy]![���Kategor], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.Formula = [nachisleniy]![Formula], Adding.Propis = [MainOccupant]![NLODGERF], Adding.NameKat = [nachisleniy]![Kategor] WHERE (((Adding.KodDoc)=" + Str(n) + "))")
RS.MoveNext
                                       Loop
 
 
 '��������� �����
'1.���� ������ ����������� �������
'mconn.Execute ("INSERT INTO Adding ( KodKv, DataR, KodN, NameN, KodDoc, ispr, com, tip ) SELECT Doc.KodKv, Doc.DataR, Doc.KodN, Doc.NameN, Doc.Key, 0, doc.com, doc.tip  From Doc WHERE (((Doc.Cod)=" + Kod + ") and (doc.stst=0))")
'2.���� ���� ����������� �������
'mconn.Execute ("INSERT INTO Adding ( KodKv, DataR, KodN, NameN, KodDoc, SummaI, ispr, com, tip ) SELECT Doc.KodKv, Doc.DataR, Doc.KodN, Doc.NameN, Doc.Key, Doc.Summa, doc.stst, doc.com, doc.tip From Doc WHERE (((Doc.Cod)=" + Kod + ") and (doc.stst=1))")

'��������� ��������� ������ ���� Adding
'If Not rs_Adding.EOF Then rs_Adding.MoveFirst
                               'Do While Not rs_Adding.EOF
                             '  For Rw = 1 To FG.Rows - 1
                             ' Nm = FG.TextMatrix(Rw, 5)
                              'Nd = FG.TextMatrix(Rw, 1)
                              
                              
'                             ��������



                            'rs_Adding.MoveNext
                                  'Next Rw
'MsgBox (FG.TextMatrix(1, 1))
Mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 1 WHERE (((ReestrDoc.Cod)=" + Fg.TextMatrix(1, 1) + "))")

'If MsgBox("������ �� ��������� ���������. ����������� ������� ����� �����������������, �������� � ������ ��������", vbYesNo) = vbYes Then
Pod.Label1.Caption = "������ ��������� �������"
Pod.Command1.Visible = True

'SposobR2.Show
'End If

End Sub

Private Sub �������_Click()
'mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")
Dim DelItem As String
Dim DelNum As String
With RS
DelItem = Fg.TextMatrix(Fg.Row, 8)
DelNum = Fg.TextMatrix(Fg.Row, 5)
If MsgBox("�� ������ ������� ���������� " + Fg.TextMatrix(Fg.Row, 4) + " ��� " + Fg.TextMatrix(Fg.Row, 6) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If RS("Key") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
.UpdateBatch
Fg.DataRefresh
'MsgBox (DelItem)

Mconn.Execute ("DELETE Adding.KodDoc, Adding.* From Adding WHERE (((Adding.KodDoc)= " + DelItem + "))")
'On Error Resume Next
If Fg.Rows <> 1 Then
Mconn.Execute ("UPDATE ReestrDoc SET ReestrDoc.Status = 0 WHERE (((ReestrDoc.Cod)=" + Fg.TextMatrix(1, 1) + "))")
If .EOF Then .MoveLast
End If


End If
End With

'������ ������ �� ������
MainForm.RSaldoN DelNum
'������ ������ � ����������
MainForm.���������������� DelNum
MainForm.RSaldoK DelNum

End Sub

Private Sub ��������()




 '��� ����� Recordset ��� Combo �������, ������������ �� ����������
 '������ � ����� ���������
On Error GoTo l3
 Rs_Combo1.Close
l3:
sq1 = "SELECT MainOccupant.Numer,MainOccupant.FAM,MainOccupant.IM,MainOccupant.OT, MainOccupant.kv_num, MainOccupant!FAM+" & Chr(34) & " " & Chr(34) + "+MainOccupant!IM+" + Chr(34) + " " + Chr(34) + " + MainOccupant!OT " + " AS ���, "
'MsgBox (sq1)
sq1 = sq1 & Chr(34) & "��." & Chr(34) & "+MainOccupant.Kv_Num+" & Chr(34)
sq1 = sq1 + "��� � " & Chr(34) & "+KLS_PODR!Num AS ���, MainOccupant.Dom FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom=KLS_PODR.���"
Kod1 = ""

If Val(Combo3.Text) <> 0 Then

'Kod = Str(Val(Left(Combo3.Text, 5)))

Kod = Str(Val(Left(Combo3.Text, InStr(1, Combo3.Text, " ", vbTextCompare))))
'KOd1 = InStr(1, Combo3.Text, " ", vbTextCompare)

'MsgBox (Kod)
sq1 = sq1 + " WHERE (((MainOccupant.Dom)=" + Kod + ")) ORDER BY MainOccupant.FAM"
End If

Rs_Combo1.Open (sq1)


End Sub
Private Sub ����()
Dim s As Double
Dim Kol As Integer

s = 0
Kol = 0
For rw = 1 To Fg.Rows - 1
If Fg.TextMatrix(rw, 7) <> "" Then s = s + Round(Fg.TextMatrix(rw, 7), 2)
Kol = Kol + 1
Next rw
Label5 = Str(Round(s, 2))
Label6 = Str(Kol)
End Sub
Private Sub ��������()
ComboQ = "Where(((Adding.KodKv) = " & Nm & "))"
ComboQ = "Where(((Adding.KodKv) = " & Nm & " and (adding.Koddoc)= " + Nd + "))"
'mconn.Execute ("UPDATE Adding INNER JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.NameN = [Nachisleniy]![Naim], Adding.KodKat = [Nachisleniy]![���Kategor], Adding.Formula = [Nachisleniy]![Formula], Adding.Tip = [Nachisleniy]![Tip], Adding.NameKat = [Nachisleniy]![Kategor] " + ComboQ)
' ���� �������
'mconn.Execute ("UPDATE Settings, Adding SET  Adding.DataR = [Settings]![TekData]" + ComboQ)
'������

'mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.TipKvKod = [MainOccupant]![KV], Adding.TipDomKod = [MainOccupant]![DomTip]" + �omboQ)
'����������
'mconn.Execute ("UPDATE Adding SET Adding.Socmin =0 " + ComboQ)
'mconn.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.Propis = Socmin.koli) AND (Adding.KodKat = Socmin.KodKategor) SET Adding.Socmin = [Socmin]![Value]" + ComboQ)
'�����
'mconn.Execute ("UPDATE Adding SET Adding.Tarif = 0 " + ComboQ)
'mconn.Execute ("UPDATE Adding INNER JOIN Tarif ON (Tarif.KodDOM = Adding.TipDomKod) AND (Tarif.KodKV = Adding.TipKvKod) AND (Adding.KodKat = Tarif.KodKat) SET Adding.Tarif = [Tarif]![Value]" + ComboQ)
'������
'��������� ������ ���������� 0 ���� ������ ����������� �������
'mconn.Execute ("UPDATE Adding SET Adding.ispr = 0 WHERE (((Adding.ispr)<>1) and ((Adding.KodKv) = " & Nm & ") and ((adding.Koddoc)= " + Nd + "))")

Mconn.Execute ("UPDATE Adding LEFT JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.LgotaVid = [Nachisleniy]![Vid]" + ComboQ)

'Rs.Requery
End Sub

Private Sub ����()
Dim rw As Integer
For rw = 1 To Fg.Rows - 1
'MsgBox (fg1.TextMatrix(fw, 27))
If Fg.TextMatrix(rw, 10) = 1 Then
'FG1.Cell(flexcpFontBold, Rw, 1, Rw, FG1.Cols) = True
'fg1.Cell(flexcpBackColor, Rw, 0) = vbCyan
Fg.Cell(flexcpFontBold, rw, 7, rw, 7) = True
Fg.Cell(flexcpBackColor, rw, 7, rw, 7) = vbCyan
End If
'If fg1.TextMatrix(Rw, 23) = "+" Then fg1.Cell(flexcpForeColor, Rw, 18, Rw, 18) = vbBlue
'If fg1.TextMatrix(Rw, 23) = "-" Then fg1.Cell(flexcpForeColor, Rw, 18, Rw, 18) = vbRed
Next rw
'fg1.Refresh

End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

 'Doc.Enabled = False
 
 If Fg.TextMatrix(0, Fg.Col) = "����������" Then
  Doc.Enabled = False
 EditCom.Show
  End If
  
  If Fg.TextMatrix(0, Fg.Col) = "...." Then
                      If Fg.TextMatrix(Fg.Row, 7) = "0" Or Fg.TextMatrix(Fg.Row, 3) = "-1" Or Fg.TextMatrix(Fg.Row, 6) = "........." Then
  MsgBox ("������� �� ��� ��������� ����������� ��� �������")
  Else
  'Doc.Enabled = False
  'Razn.Show
  Filter.Nm = Fg.TextMatrix(Fg.Row, 5)
  
  MainForm.Fnum = Fg.TextMatrix(Fg.Row, 5)
  Lic.a = "Doc"
  Lic.Show
                                     End If
  
  End If
 End Sub
Private Sub �������()

For rw = 1 To Fg.Rows - 1
If Fg.TextMatrix(rw, 9) <> "" Then
If Fg.TextMatrix(rw, 9) = 1 Then
Fg.Cell(flexcpForeColor, rw, 1, rw, 10) = vbBlue
Fg.Cell(flexcpFontBold, rw, 1, rw, 10) = True


End If
End If

If InStr(Fg.TextMatrix(rw, 6), "�/C") <> 0 Then

Fg.Cell(flexcpForeColor, rw, 1, rw, 10) = vbRed
Fg.Cell(flexcpFontBold, rw, 1, rw, 10) = True

End If
Next
End Sub

Function XCol_(ByVal Column_ As Long) As String
    If (Column_ < 0) Then Column_ = 0
    If (Column_ < 26) Then
        XCol_ = Chr(Column_ + Asc("A"))
    ElseIf (Column_ < 676) Then
        XCol_ = Chr((Column_ \ 26) + Asc("A") - 1) & Chr((Column_ Mod 26) + Asc("A"))
    Else
        XCol_ = "ZZ"
    End If
End Function

'****************


'******************

 Function sumPropis(dSumma As Double) As String
Dim Sp As String ' ������ ��������
Dim sn As String ' �������� ������������� �����
Dim sd As String ' ���������� �������
Dim rub(10) As String ' ����� ������
Dim mlrd(10) As String ' ����� ����������
Dim mln(10) As String ' ����� ���������
Dim tys(10) As String ' ����� �����

rub(1) = " ����� "
rub(2) = " ����� "
rub(3) = " ����� "
rub(4) = " ����� "
rub(5) = " ������ "
rub(6) = " ������ "
rub(7) = " ������ "
rub(8) = " ������ "
rub(9) = " ������ "
rub(0) = " ������ "
'
tys(1) = " ������ "
tys(2) = " ������ "
tys(3) = " ������ "
tys(4) = " ������ "
tys(5) = " ����� "
tys(6) = " ����� "
tys(7) = " ����� "
tys(8) = " ����� "
tys(9) = " ����� "
tys(0) = " ����� "
'
mln(1) = " ������� "
mln(2) = " �������� "
mln(3) = " �������� "
mln(4) = " �������� "
mln(5) = " ��������� "
mln(6) = " ��������� "
mln(7) = " ��������� "
mln(8) = " ��������� "
mln(9) = " ��������� "
mln(0) = " ��������� "
'
mlrd(1) = " �������� "
mlrd(2) = " ��������� "
mlrd(3) = " ��������� "
mlrd(4) = " ��������� "
mlrd(5) = " ���������� "
mlrd(6) = " ���������� "
mlrd(7) = " ���������� "
mlrd(8) = " ���������� "
mlrd(9) = " ���������� "
mlrd(0) = " ���������� "
'
'�������������
Let sumPropis = ""
'��������� ����� �� ������������
If dSumma <= 0 Then Exit Function
'��������� �� �������
sn = Format(Int(dSumma), "000000000000")
sd = Format(Round((dSumma - Val(sn)) * 100, 0), "00")
'���������������� ������
'��������� - ����� ����� ����������
If Val(Mid(sn, 1, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 1, 3), False) & IIf(Mid(sn, 2, 1) = 1, mlrd(0), mlrd(Mid(sn, 3, 1)))
'��������
If Val(Mid(sn, 4, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 4, 3), False) & IIf(Mid(sn, 5, 1) = 1, mln(0), mln(Mid(sn, 6, 1)))
'������
If Val(Mid(sn, 7, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 7, 3), True) & IIf(Mid(sn, 8, 1) = 1, tys(0), tys(Mid(sn, 9, 1)))
'� �������
sumPropis = sumPropis & sTriple(Mid(sn, 10, 3), False)
'���������� ���������
sumPropis = sumPropis & IIf(Mid(sn, 11, 1) = 1, rub(0), rub(Right(sn, 1))) & sd & " ���."
'
End Function

Function sTriple(sRazr As String, bGender As Boolean) As String
'������� ��������� ����������� ����� � ����� �������� � ������ ����
Dim Ed(20) As String  ' ������ ������
Dim des(10) As String ' ������ �������
Dim sot(10) As String ' ������ �����
'�������� ������
Ed(0) = ""
Ed(1) = " ����"
Ed(2) = " ���"
Ed(3) = " ���"
Ed(4) = " ������"
Ed(5) = " ����"
Ed(6) = " �����"
Ed(7) = " ����"
Ed(8) = " ������"
Ed(9) = " ������"
Ed(10) = " ������"
Ed(11) = " �����������"
Ed(12) = " ����������"
Ed(13) = " ����������"
Ed(14) = " ������������"
Ed(15) = " ����������"
Ed(16) = " �����������"
Ed(17) = " ����������"
Ed(18) = " ������������"
Ed(19) = " ������������"
'�������� ��������
des(0) = ""
des(1) = " ������"
des(2) = " ��������"
des(3) = " ��������"
des(4) = " �����"
des(5) = " ���������"
des(6) = " ����������"
des(7) = " ���������"
des(8) = " �����������"
des(9) = " ���������"
'�������� �����
sot(0) = ""
sot(1) = " ���"
sot(2) = " ������"
sot(3) = " ������"
sot(4) = " ���������"
sot(5) = " �������"
sot(6) = " ��������"
sot(7) = " �������"
sot(8) = " ���������"
sot(9) = " ���������"
' ���� ���� ��� �����
If bGender Then
    Ed(1) = " ����"
    Ed(2) = " ���"
End If
' ���������� � �������
sTriple = sTriple & sot(Mid(sRazr, 1, 1))
' ���� ������� �������
If Mid(sRazr, 2, 2) > 10 And Mid(sRazr, 2, 2) < 20 Then
    sTriple = sTriple & Ed(Mid(sRazr, 2, 2))
Else
' ����� ������ - ���� ������� �� ������
    sTriple = sTriple & des(Mid(sRazr, 2, 1))
    sTriple = sTriple & Ed(Mid(sRazr, 3, 1))
End If

End Function




