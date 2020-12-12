VERSION 5.00
Begin VB.Form ImpLg 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Импорт льгот"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form8"
   ScaleHeight     =   7410
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
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
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   9135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "Путь к импортируемим файлам"
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
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   9135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5640
      X2              =   9600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   9600
      X2              =   9600
      Y1              =   1320
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   9600
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   4200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Шаг № 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Окно импорта данных о льготах на основании данных предоставляемых органами соцзащиты"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "ImpLg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FileImp.Show
End Sub

Private Sub Command2_Click()
Dim Cn As ADODB.Connection
Dim Gek As ADODB.Recordset
Dim Baza As ADODB.Recordset

Set Cn = New ADODB.Connection

Cn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы dBASE;Initial Catalog= Import\"

Set Gek = New ADODB.Recordset
Set Baza = New ADODB.Recordset
Gek.Open ("SELECT GEK.* FROM GEK"), Cn
Baza.Open ("Baza"), Cn



Set Cn = Nothing
End Sub

Private Sub Form_Load()
Dim cnA As ADODB.Connection





Set cnA = New ADODB.Connection


cnA.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True;Persist Security Info=True"
cnA.Open "data/Kvartplata.mdb"




'cnA.Execute ("SELECT GEK2.* INTO GEK FROM GEK2")
End Sub
