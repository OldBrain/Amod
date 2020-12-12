VERSION 5.00
Begin VB.Form Information 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6345
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9945
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   46
      Text            =   "Text13"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9240
      TabIndex        =   44
      Text            =   "Text12"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1560
      TabIndex        =   42
      Text            =   "Text11"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   6600
      TabIndex        =   41
      Text            =   "Text10"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5640
      TabIndex        =   40
      Text            =   "Text9"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4800
      TabIndex        =   39
      Text            =   "Text8"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3840
      TabIndex        =   38
      Text            =   "Text7"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   37
      Text            =   "Text6"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8040
      TabIndex        =   36
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5040
      TabIndex        =   35
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   34
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   33
      Top             =   1680
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   32
      Top             =   2040
      Width           =   8535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Телефон"
      Height          =   255
      Left            =   8040
      TabIndex        =   45
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Кол-во комнат"
      Height          =   255
      Left            =   8040
      TabIndex        =   43
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label42"
      Height          =   255
      Left            =   7560
      TabIndex        =   31
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Документ на квартиру"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Паспорт"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "по -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   28
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Дата прописки с"
      Height          =   255
      Left            =   3480
      TabIndex        =   27
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Дата рождения"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   7920
      X2              =   7920
      Y1              =   4320
      Y2              =   5640
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   5640
      Y2              =   4800
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   4800
      Y2              =   5640
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   5640
      Y2              =   4800
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   3720
      X2              =   3720
      Y1              =   5640
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   2760
      X2              =   2760
      Y1              =   4800
      Y2              =   5640
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4320
      Y2              =   5640
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1440
      X2              =   1440
      Y1              =   4800
      Y2              =   5640
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Коридор/холл"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Балкон"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Общая"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Полезная"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Кухня"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Ванная"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Туалет"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label20"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label27"
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
      Left            =   7680
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label26"
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
      Left            =   7680
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label25"
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
      Left            =   8880
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Этаж"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   7815
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   7920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Прописано"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Проживает"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Тип квартиры"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Тип дома"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   8655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Адрес"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   8655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
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
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ответственный квартиросъемщик"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   9735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Лицевой счет №"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Zap
If Text3 <> "" Then Filter.infRS.Fields("BIRTHDAY") = Text3
If Text4 <> "" Then Filter.infRS.Fields("LDATEBEG") = Text4
If Text5 <> "" Then Filter.infRS.Fields("LDATEEND") = Text5
Filter.infRS.Fields("PASSPORT") = Text1
If Text2 <> "" Then Filter.infRS.Fields("NORDER") = Text2
Filter.infRS.Fields("HABSPACE") = Text11
Filter.infRS.Fields("KITCHSPACE") = Text6
Filter.infRS.Fields("BATHSPACE") = Text7
Filter.infRS.Fields("TOILSPACE") = Text8
Filter.infRS.Fields("BALCSPACE") = Text9
Filter.infRS.Fields("CORRSPACE") = Text10
Filter.infRS.Fields("NROOM") = Text12
Filter.infRS.Fields("TELEPHONE") = Text13
Filter.infRS.UpdateBatch
Unload Me
Exit Sub
Zap:
Filter.infRS.UpdateBatch
Unload Me
End Sub

Private Sub Form_Load()
Filter.Enabled = False


Label42.Caption = Filter.infRS("Numer")
Label2.Caption = Filter.infRS("OldNum")
Label4.Caption = Filter.infRS("Fam") + " " + Filter.infRS("Im") + " " + Filter.infRS("Ot")
Label6.Caption = Filter.FG.Cell(flexcpText, Filter.FG.Row, 5) + " дом №" + Filter.FG.Cell(flexcpText, Filter.FG.Row, 6) + " Кв.№ " + Filter.FG.Cell(flexcpText, Filter.FG.Row, 9)
Label8.Caption = Filter.infRS("Name_Dom")
Label9.Caption = Filter.infRS("Name_Kv")

Label20.Caption = Filter.infRS("COMSPACE")
Text11 = Filter.infRS("HABSPACE")
Text6 = Filter.infRS("KITCHSPACE")
Text7 = Filter.infRS("BATHSPACE")
Text8 = Filter.infRS("TOILSPACE")
Text9 = Filter.infRS("BALCSPACE")


Text12 = Filter.infRS("NROOM")
Text13 = Filter.infRS("TELEPHONE")

'TELEPHONE

Text10 = Filter.infRS("CORRSPACE")

Label26.Caption = Filter.infRS("NLODLIFT")
Label27.Caption = Filter.infRS("NLODGERF")

Label25.Caption = Filter.infRS("floor")

If Filter.infRS("BIRTHDAY") <> "" Then Text3 = Filter.infRS("BIRTHDAY")

If Filter.infRS("LDATEBEG") <> "" Then Text4 = Filter.infRS("LDATEBEG")
If Filter.infRS("LDATEEND") <> "" Then Text5 = Filter.infRS("LDATEEND")

If Filter.infRS("PASSPORT") <> "" Then Text1 = Filter.infRS("PASSPORT")

If Filter.infRS("NORDER") <> "" Then Text2 = Filter.infRS("NORDER")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True
End Sub
