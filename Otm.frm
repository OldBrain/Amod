VERSION 5.00
Begin VB.Form Otm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1788
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   4896
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Отмена"
      Height          =   615
      Left            =   120
      Picture         =   "Otm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Приметить"
      Height          =   615
      Left            =   2400
      Picture         =   "Otm.frx":007B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Text            =   "Адрес"
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Фильтр по адресу"
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
      Left            =   480
      TabIndex        =   3
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4410
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
      Left            =   0
      Picture         =   "Otm.frx":01F9
      Top             =   0
      Width           =   156
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "Otm.frx":0443
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "Otm.frx":0B8D
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   360
      Picture         =   "Otm.frx":12D7
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Otm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim mconn As ADODB.Connection
Dim Addrconn As ADODB.Recordset
Dim D As String



Private Sub Combo1_Change()
'AddrConn.MoveFirst
'Do While Not AddrConn.EOF
'If Combo1.Text = AddrConn("NAIM_KLS") + " дом № " + AddrConn("Num") Then
'D = AddrConn("Код")
'AddrConn("DOMTip") = AddrConn("Tip")
'End If
'AddrConn.MoveNext
'Loop




End Sub

Private Sub Command1_Click()
If Combo1.Text = "Адрес" Then Exit Sub

Filter.FG.Enabled = True
Filter.FG.Row = 1
Filter.FG.TextMatrix(1, 5) = Combo1.Text
SendKeys "{Enter}"
Unload Me
End Sub

Private Sub Command2_Click()
Filter.FG.Row = Filter.oldR
Filter.Enabled = True
Unload Me
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
MakeWindow Me, False
Me.KeyPreview = True

Filter.Enabled = False

' open connection
'  Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
  
  
  Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic


'Addrconn.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim FROM KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Addrconn.Open ("SELECT KLS_PODR.NAIM_KLS FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД GROUP BY KLS_PODR.NAIM_KLS ORDER BY KLS_PODR.NAIM_KLS")

Addrconn.MoveFirst
Do While Not Addrconn.EOF

Combo1.AddItem Addrconn("NAIM_KLS")
Addrconn.MoveNext
Loop

SendKeys "{F4}"
End Sub
Private Sub Пометить()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True
Filter.FG.SetFocus
Filter.Command15.Visible = True
Filter.Command17.Visible = True
End Sub
