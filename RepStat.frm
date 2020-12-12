VERSION 5.00
Begin VB.Form RepStat 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3510
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "RepStat.frx":0000
   LinkTopic       =   "Статистические отчеты"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   StartUpPosition =   2  'CenterScreen
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      Caption         =   "Список льготников"
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Сведения о приватизации"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "По видам удобств"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   4455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Излишки"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Количество прописанных, метраж"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Запрос Количествопроп"
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "АРМ ""Квартплата + "" Статистика"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4170
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   360
      Picture         =   "RepStat.frx":030A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   600
      Picture         =   "RepStat.frx":0A54
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   840
      Picture         =   "RepStat.frx":119E
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
      Height          =   240
      Left            =   0
      Picture         =   "RepStat.frx":18E8
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "RepStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Reports
Unload RepStat
Analizlgot.G = 8
Reports.sq = "Количествопроп"
Analizlgot.Об 3
Unload Me
Analizlgot.Show
End Sub

Private Sub Command2_Click()
'Unload Me
'RepLgota.Show
End Sub

Private Sub Command3_Click()
Unload Me
Reports.Enabled = True
End Sub

Private Sub Command4_Click()
If MsgBox("Разбивать по подразделениям", vbYesNo) = vbYes Then
Reports.sq = "SELECT KLS_PODR.Подразделение, KLS_PODR.Благ, TipDom.Name_Dom, TipKv.Name_Kv, Count(MainOccupant.Numer) AS [Кол квартир], Sum(MainOccupant.COMSPACE) AS [Общ площадь], Sum(MainOccupant.HABSPACE) AS [Полезная площадь], Sum(MainOccupant.NLODGERF) AS Прописано, Sum(MainOccupant.NLODGER) AS Проживает, Sum(MainOccupant.NROOM) AS [Кол-во комнат], Sum(MainOccupant.KITCHSPACE) AS Кухня, Sum(MainOccupant.BATHSPACE) AS Вынная, Sum(MainOccupant.CORRSPACE) AS Коридор, Sum(MainOccupant.TOILSPACE) AS Туалет, Sum(MainOccupant.BALCSPACE) AS Балкон FROM KLS_PODR INNER JOIN ((MainOccupant INNER JOIN TipKv ON MainOccupant.KV = TipKv.Код) INNER JOIN TipDom ON MainOccupant.DomTip = TipDom.Код) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.Подразделение, KLS_PODR.Благ, TipDom.Name_Dom, TipKv.Name_Kv ORDER BY TipDom.Name_Dom, TipKv.Name_Kv"
Analizlgot.G = 16
Else
Reports.sq = "SELECT KLS_PODR.Благ, TipDom.Name_Dom, TipKv.Name_Kv, Count(MainOccupant.Numer) AS [Кол квартир], Sum(MainOccupant.COMSPACE) AS [Общ площадь], Sum(MainOccupant.HABSPACE) AS [Полезная площадь], Sum(MainOccupant.NLODGERF) AS Прописано, Sum(MainOccupant.NLODGER) AS Проживает, Sum(MainOccupant.NROOM) AS [Кол-во комнат], Sum(MainOccupant.KITCHSPACE) AS Кухня, Sum(MainOccupant.BATHSPACE) AS Вынная, Sum(MainOccupant.CORRSPACE) AS Коридор, Sum(MainOccupant.TOILSPACE) AS Туалет, Sum(MainOccupant.BALCSPACE) AS Балкон FROM KLS_PODR INNER JOIN ((MainOccupant INNER JOIN TipKv ON MainOccupant.KV = TipKv.Код) INNER JOIN TipDom ON MainOccupant.DomTip = TipDom.Код) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.Благ, TipDom.Name_Dom, TipKv.Name_Kv ORDER BY TipDom.Name_Dom, TipKv.Name_Kv"
Analizlgot.G = 15
End If



'AnalizLgot.G = 16
'Reports.sq = "SELECT KLS_PODR.Подразделение, KLS_PODR.Благ, TipDom.Name_Dom, TipKv.Name_Kv, Count(MainOccupant.Numer) AS [Кол квартир], Sum(MainOccupant.COMSPACE) AS [Общ площадь], Sum(MainOccupant.HABSPACE) AS [Полезная площадь], Sum(MainOccupant.NLODGERF) AS Прописано, Sum(MainOccupant.NLODGER) AS Проживает, Sum(MainOccupant.NROOM) AS [Кол-во комнат], Sum(MainOccupant.KITCHSPACE) AS Кухня, Sum(MainOccupant.BATHSPACE) AS Вынная, Sum(MainOccupant.CORRSPACE) AS Коридор, Sum(MainOccupant.TOILSPACE) AS Туалет, Sum(MainOccupant.BALCSPACE) AS Балкон FROM KLS_PODR INNER JOIN ((MainOccupant INNER JOIN TipKv ON MainOccupant.KV = TipKv.Код) INNER JOIN TipDom ON MainOccupant.DomTip = TipDom.Код) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.Подразделение, KLS_PODR.Благ, TipDom.Name_Dom, TipKv.Name_Kv ORDER BY TipDom.Name_Dom, TipKv.Name_Kv"
Analizlgot.Об 3
Unload Me
Analizlgot.Show
End Sub

Private Sub Command5_Click()
Unload Reports
Unload RepStat
Analizlgot.G = 11
Reports.sq = "SELECT [MainOccupant]![Priv] AS [ПРИВАТИЗИРОВАНА(Да/Нет)], KLS_PODR.NAIM_KLS as [Адрес],  MainOccupant.LDOK as [Документ на квартиру], MainOccupant.FAM as [Фамилия], MainOccupant.IM as [Имя], MainOccupant.OT as [Отчество], MainOccupant.kv_num as [№ кв], MainOccupant.NLODGERF as [Прописано], MainOccupant.COMSPACE as [Общ площадь] FROM KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom ORDER BY [MainOccupant]![Priv], KLS_PODR.NAIM_KLS"

Analizlgot.Об 2
Unload Me
Analizlgot.Show
End Sub

Private Sub Command7_Click()
Rep_Izl.Show
Unload Me
End Sub

Private Sub Form_Load()
Reports.Enabled = False
MakeWindow Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Reports.Enabled = True

End Sub

Private Sub xpcmdbutton1_Click()
Unload Reports
Unload RepStat
Analizlgot.Titl = "Список льготников " + MainMenu.Command13.Caption
Analizlgot.G = 12
Reports.sq = "SELECT Lgota.NAME_KLS AS [Категория льготы], KLS_PODR.NAIM_KLS AS Улица, KLS_PODR.Num AS Дом, MainOccupant.kv_num AS [Кв №], MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.NLODGERF AS Прописано, MainOccupant.NLODGER AS Проживает, MainOccupant.COMSPACE AS Площадь, MainOccupant.Priv AS Приватизация FROM Lgota INNER JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Lgota.NomNum = MainOccupant.Numer ORDER BY Lgota.NAME_KLS, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.kv_num"
Analizlgot.Об 0
Unload Me
Analizlgot.Show
End Sub
