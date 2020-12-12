VERSION 5.00
Begin VB.Form ShapkaTBL 
   Caption         =   "Меню настройки документа"
   ClientHeight    =   5604
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   8688
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   5604
   ScaleWidth      =   8688
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   6720
      TabIndex        =   21
      Text            =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   3000
      TabIndex        =   17
      Text            =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   5160
      TabIndex        =   16
      Text            =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Расчет ОДН"
      Height          =   252
      Left            =   7200
      TabIndex        =   15
      Top             =   720
      Width           =   1452
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1440
      TabIndex        =   13
      Text            =   "*"
      Top             =   1800
      Width           =   1572
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Простои"
      Height          =   252
      Left            =   7200
      TabIndex        =   12
      Top             =   240
      Width           =   1692
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "0"
      Top             =   4560
      Width           =   8415
   End
   Begin VB.TextBox Text1 
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
      Left            =   120
      TabIndex        =   8
      Text            =   " "
      Top             =   3480
      Width           =   8415
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   7
      Text            =   "Combo3"
      Top             =   1200
      Width           =   5655
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   720
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "ShapkaTBL.frx":0000
      Left            =   2880
      List            =   "ShapkaTBL.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Создать документ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "ДАННЫЕ ОБЩЕДОМОВОГО СЧЕТЧИКА"
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   4800
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label8 
      Caption         =   "Сумарные данные счетчиков"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.Label Label7 
      Caption         =   "Общая площадь дома"
      Height          =   372
      Left            =   3000
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.Label Label6 
      Caption         =   "Подъезд:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"ShapkaTBL.frx":0004
      Height          =   492
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   8412
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Коментарий"
      Height          =   372
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   8532
   End
   Begin VB.Label Label3 
      Caption         =   "Адрес:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Начисление:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Категория расчета:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "ShapkaTBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cRs As ADODB.Recordset
Dim odnRs As ADODB.Recordset
Public Kat As Integer
Public KNac As Integer
Public KAdres As Integer
Public ODN_Plo As Integer
Public ODN_Sc As Double
Public Prostoy As Boolean
Public Odn As Boolean

Private Sub Check1_Click()

If Me.Check1.Value = 1 Then
Me.Check2.Value = 0
Me.Combo4.Enabled = True
Me.Prostoy = True
Me.Text2 = "round((((S1*S2)/S3)*S4)*-1,2)"
Else
Me.Combo4.Enabled = False
Prostoy = False
PodD = "*"
Me.Check2.Value = 0
End If
podezd
End Sub

Private Sub Check2_Click()
Me.Check1.Value = 0

If Me.Check2.Value = 1 Then
Me.Text3.Visible = True
Me.Text4.Visible = True
Me.Label7.Visible = True
Me.Label8.Visible = True
Me.Label9.Visible = True
Me.Text5.Visible = True
Me.Odn = True
Me.Text2 = "round((S4/S3)*S2*S1,2)"
Else
Me.Text3.Visible = False
Me.Text4.Visible = False
Me.Label7.Visible = False
Me.Label8.Visible = False

Me.Label9.Visible = False
Me.Text5.Visible = False
Odn = False
Me.Text2 = "0"
End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
'MsgBox "211"
Kat = Val(Combo1.Text)
If Kat <> 0 Then Combo2.Enabled = True
ReestrTablDoc.TM.Open ("SELECT nachisleniy.КодKategor, nachisleniy.Kod, nachisleniy.Naim From Nachisleniy WHERE (((nachisleniy.КодKategor)=" + Str(Kat) + "))")
If ReestrTablDoc.TM.EOF = False Then ReestrTablDoc.TM.MoveFirst Else Combo2.Enabled = False
If ReestrTablDoc.TM.EOF = False Then Combo2.Text = Str(ReestrTablDoc.TM("Kod")) + " " + ReestrTablDoc.TM("Naim") Else MsgBox "Нет начислений по данной категории расчета"
Do While Not ReestrTablDoc.TM.EOF
Combo2.AddItem Str(ReestrTablDoc.TM("Kod")) + " " + ReestrTablDoc.TM("Naim")
ReestrTablDoc.TM.MoveNext
Loop
ReestrTablDoc.TM.Close
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
podezd
End Sub

Private Sub Command1_Click()
b = Combo3.Text
a = InStr(b, "|")
b = Left(b, a - 1)
KAdres = Val(b)
'Val (Combo3.Text)
KNac = Val(Combo2.Text)
If Me.Prostoy = False Then PodD = Val(Combo4.Text) Else PodD = "*"
ReestrTablDoc.НовыйТБЛ
VibTablDoc.Show
'MsgBox Str(KAdres) + "  " + Str(KNac)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo2.Enabled = False

ReestrTablDoc.TM.Open ("SELECT Kategor.Код, Kategor.Name_Kategor FROM Kategor")
ReestrTablDoc.TM.MoveFirst
Combo1.Text = Str(ReestrTablDoc.TM("Код")) + " " + ReestrTablDoc.TM("Name_Kategor")
Do While Not ReestrTablDoc.TM.EOF
Combo1.AddItem Str(ReestrTablDoc.TM("Код")) + " " + ReestrTablDoc.TM("Name_Kategor")
ReestrTablDoc.TM.MoveNext
Loop
ReestrTablDoc.TM.Close


ReestrTablDoc.TM.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM KLS_PODR")
ReestrTablDoc.TM.MoveFirst
Combo3.Text = Str(ReestrTablDoc.TM("Код")) + "|" + ReestrTablDoc.TM("NAIM_KLS") + " Дом №" + ReestrTablDoc.TM("Num")
Do While Not ReestrTablDoc.TM.EOF
Combo3.AddItem Str(ReestrTablDoc.TM("Код")) + "|" + ReestrTablDoc.TM("NAIM_KLS") + " Дом №" + ReestrTablDoc.TM("Num")
ReestrTablDoc.TM.MoveNext
Loop
ReestrTablDoc.TM.Close
End Sub



Private Sub podezd()
Me.Combo4.Clear
Me.Combo4.Text = "*"
Me.Combo4.AddItem "*"
b = Combo3.Text
a = InStr(b, "|")
b = Left(b, a - 1)
KAdres = Val(b)
If Me.Prostoy = True Then


Set cRs = New ADODB.Recordset
If KAdres <> 0 Then
cRs.Open ("SELECT MainOccupant.podyezd, MainOccupant.Dom From MainOccupant GROUP BY MainOccupant.podyezd, MainOccupant.Dom HAVING (((MainOccupant.Dom)=" + Str(KAdres) + "))"), Mconn
Else
cRs.Open ("SELECT MainOccupant.podyezd From MainOccupant GROUP BY MainOccupant.podyezd"), Mconn
End If
Do While Not cRs.EOF
'If cRs("podyezd") <> "" Then
Me.Combo4.AddItem cRs("podyezd")
cRs.MoveNext
Loop
End If

If Me.Check2.Value = 1 Then
'считаем общую площадь и счетчики
Set odnRs = New ADODB.Recordset
odnRs.Open ("SELECT MainOccupant.Dom, Sum(MainOccupant.COMSPACE) AS [Sum-COMSPACE] From MainOccupant GROUP BY MainOccupant.Dom HAVING (((MainOccupant.Dom)=" + Str(KAdres) + "))"), Mconn
ODN_Plo = odnRs("Sum-COMSPACE")
Me.Text3.Text = ODN_Plo
odnRs.Close
'считаем общие показания счетчиков
Set odnRs = New ADODB.Recordset
odnRs.Open ("SELECT MainOccupant.Dom, Adding.KodKat, Sum(Adding.Shc_new) AS [Sum-Shc_new], Sum(Adding.Shc_old) AS [Sum-Shc_old], Sum([Adding]![Shc_new]-[Adding]![Shc_old]) AS Разница FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer GROUP BY MainOccupant.Dom, Adding.KodKat HAVING (((MainOccupant.Dom)=" + Str(KAdres) + ") AND ((Adding.KodKat)=" + Str(Kat) + "))"), Mconn, adOpenStatic, adLockBatchOptimistic
'odnRs.MoveLast
ODN_Sc = 0
If odnRs.RecordCount <> 0 Then ODN_Sc = odnRs("Разница")
Me.Text4.Text = ODN_Sc
End If
End Sub

