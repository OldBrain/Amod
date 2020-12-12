VERSION 5.00
Begin VB.Form KvitShapka 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочные данные"
   ClientHeight    =   3132
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3744
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "KvitShapka.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3132
   ScaleWidth      =   3744
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Сохранить"
      Height          =   492
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отмена"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1212
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
      Height          =   1812
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "KvitShapka.frx":030A
      Top             =   240
      Width           =   3132
   End
End
Attribute VB_Name = "KvitShapka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Len(Me.Text1.Text) > 255 Then Me.Text1.Text = Left(Me.Text1.Text, 255)
Mconn.Execute ("UPDATE Settings SET Settings.Kvit = '" + Me.Text1.Text + "'")
Unload Me
End Sub

Private Sub Form_Load()
Dim RsRec1 As ADODB.Recordset
'Получаем данные SETTING
Set RsRec1 = New ADODB.Recordset
'Set RsRec1.ActiveConnection = Mconn
RsRec1.Open ("SELECT Settings.Kvit FROM Settings"), Mconn
Me.Text1.Text = RsRec1("Kvit")
RsRec1.Close

End Sub

