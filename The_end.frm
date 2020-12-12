VERSION 5.00
Begin VB.Form The_end 
   Caption         =   "Выгрузка данных "
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
      Caption         =   "Выход"
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
      Caption         =   "Приватизация"
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
      Caption         =   "Тарифы"
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
      Caption         =   "Сальдо на начало периода"
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
      Caption         =   "Абоненты ЖЭК"
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
      Caption         =   "Льготы"
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
      Caption         =   "Сверка"
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
Analizlgot.Titl = "Список лицевых счетов " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 11
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, KLS_PODR.num AS Дом,MainOccupant.kv_num AS Кв, MainOccupant.OLDNUM AS [Ключ], MainOccupant.BanKN AS [N лиц сч банк], MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.COMSPACE AS Площадь, MainOccupant.NLODGERF AS Прописано FROM MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД ORDER BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num"
'Analizlgot.Об 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(250, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton2_Click()
Analizlgot.Titl = "Сальдо на начало периода " + MainMenu.Command13.Caption + "<-> Переплата <+>-Долг"
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 4
Reports.sq = "SELECT Kategor.Name_Kategor AS [Категория расчета], Saldo_Arh.KodKV AS Ключ, Saldo_Arh.SK AS Сальдо FROM Kategor INNER JOIN Saldo_Arh ON Kategor.Код = Saldo_Arh.KodKat"
'Analizlgot.Об 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton3_Click()
Analizlgot.Titl = "Список жильцов имеющих льготы " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 7
Reports.sq = "SELECT Lgota.NomNum AS Ключ, Lgota.Numer AS [Код льготы из справочника льгот], Lgota.NAME_KLS AS Наименование, Lgota.LPKV AS Процент, Lgota.USEKV AS [Способ применения], IIf([OhteCode]=0,'Отв.лвартиросъемщик','Совм.проживающий') AS [Принадлежность льготы] From Lgota ORDER BY Lgota.NomNum"
'Analizlgot.Об 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 200, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 250, 200), vbBlack, True


'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton4_Click()
Analizlgot.Titl = "Список тарифов " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 7
Reports.sq = "SELECT Tarif.Kategor AS [Категория расчета], MainOccupant.Numer AS Ключ, Tarif.NameDOM AS [Тип дома], Tarif.NameKV AS [Тип квартиры], Tarif.Value AS Тариф FROM Tarif INNER JOIN MainOccupant ON (MainOccupant.DomTip = Tarif.KodDOM) AND (Tarif.KodKV = MainOccupant.KV) ORDER BY Tarif.Kategor"
'Analizlgot.Об 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 200, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 3, , RGB(150, 250, 200), vbBlack, True


'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 1, 3, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show




End Sub

Private Sub xpcmdbutton5_Click()
Analizlgot.Titl = "Список тарифов " + MainMenu.Command13.Caption
Analizlgot.G = 17
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, KLS_PODR.Num AS Дом, MainOccupant.kv_num AS Кв, MainOccupant.COMSPACE AS Площадь, MainOccupant.NLODGERF AS Прописано, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.BanKN AS [л/сч], Tarif.NameDOM AS [Тип дома], Tarif.NameKV AS [Тип кв], Tarif.Value AS Тариф, Lgota.NAME_KLS AS Льгота, Lgota.LPKV AS Процент, Lgota.USEKV AS [Способ прим], IIf([OhteCode] Is Not Null,IIf([OhteCode]=0,'Отв.кв.','Совм.прож'),' ') AS [Принадлежность льготы] FROM ((MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) LEFT JOIN Tarif ON (MainOccupant.DomTip = Tarif.KodDOM) AND (MainOccupant.KV = Tarif.KodKV)) LEFT JOIN Lgota ON MainOccupant.Numer = Lgota.NomNum Where (((Tarif.KodKat) = 1)) ORDER BY KLS_PODR.NAIM_KLS"
Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Unload Me
Analizlgot.Show
End Sub

Private Sub xpcmdbutton6_Click()
Analizlgot.Titl = "Сведения о приватизации и льготах " + MainMenu.Command13.Caption
Analizlgot.G = 10
Reports.sq = "SELECT MainOccupant.BanKN AS Номер, KLS_PODR.NAIM_KLS AS Улица, KLS_PODR.Num AS дом, MainOccupant.kv_num AS Кв, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Lgota.NAME_KLS AS Льгота, MainOccupant.Priv AS Приватизировано FROM (Lgota RIGHT JOIN MainOccupant ON Lgota.NomNum = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД ORDER BY KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.kv_num, MainOccupant.FAM"
Analizlgot.FG1.OutlineBar = flexOutlineBarComplete
Unload Me
Analizlgot.Show



End Sub
