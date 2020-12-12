VERSION 5.00
Begin VB.Form Rep_Izl 
   Caption         =   "Выбор параметров отчета"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form8"
   ScaleHeight     =   4785
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check12 
      Caption         =   "Показывать итог нач излишек"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Показывать метраж излишек с учетом 10 м"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Показывать дополн. 10 м к СМ"
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Показывать излишки без учета 10 м"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Показывать соцминимум"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Показывать кол-во прописанных"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Показывать общую площадь"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Показывать начислено излишки"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Показывать тариф излишки"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Показывать начислено в пределах СМ"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Показывать тариф в пределах СМ"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Показывать адреса и фамилии"
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
      Caption         =   "ВСЕ КВАРТИРЫ"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ТОЛЬКО КВАРТИРЫ С ИЗЛИШКАМИ"
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




'SELECT KLS_PODR.NAIM_KLS, MainOccupant.FAM, Adding.Tarif AS [Тариф в пределах соцнормы], IIf(Adding!KodN=2,Adding!SummaI,0) AS [Начислено в  пределах соцнормы], [Adding]![TarifI] AS [Тариф излишки], IIf([Adding]![KodN]=3,[Adding]![SummaI],0) AS [Начислено излишки], [Adding]![ObPl] AS [Общая площадь], [Adding]![Propis] AS Прописано, [Adding]![Socmin] AS Соцминимум, IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0) AS [Излишки м*м], IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0) AS [10 к СМ], IIf([Излишки м*м]>[10 к СМ],Round([Излишки м*м]-[10 к СМ],1),0) AS [Излишки метраж], Round([Излишки метраж]*([Тариф излишки]+[Тариф в пределах соцнормы]),2) AS [Нач излишки]
'FROM (Adding LEFT JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД
'WHERE (((IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0))<>0) AND ((Adding.KodN)=3)) OR (((Adding.KodN)=2));
AnalizLgot.G = 2

If Check1.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "KLS_PODR.NAIM_KLS, MainOccupant.FAM" Else Q = Q + "KLS_PODR.NAIM_KLS, MainOccupant.FAM"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check2.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "Adding.Tarif AS [Тариф в пределах соцнормы]" Else Q = Q + "Adding.Tarif AS [Тариф в пределах соцнормы]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check3.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf(Adding!KodN=2,Adding!SummaI,0) AS [Начислено в  пределах соцнормы]" Else Q = Q + "IIf(Adding!KodN=2,Adding!SummaI,0) AS [Начислено в  пределах соцнормы]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check4.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![TarifI] AS [Тариф излишки]" Else Q = Q + "[Adding]![TarifI] AS [Тариф излишки]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check5.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf([Adding]![KodN]=3,[Adding]![SummaI],0) AS [Начислено излишки]" Else Q = Q + "IIf([Adding]![KodN]=3,[Adding]![SummaI],0) AS [Начислено излишки]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check6.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![ObPl] AS [Общая площадь]" Else Q = Q + "[Adding]![ObPl] AS [Общая площадь]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check7.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![Propis] AS Прописано" Else Q = Q + "[Adding]![Propis] AS Прописано"
AnalizLgot.G = AnalizLgot.G + 1
End If

If Check8.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "[Adding]![Socmin] AS Соцминимум" Else Q = Q + "[Adding]![Socmin] AS Соцминимум"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check9.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0) AS [Излишки м*м]" Else Q = Q + "IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0) AS [Излишки м*м]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check10.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0) AS [10 к СМ]" Else Q = Q + "IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0) AS [10 к СМ]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check11.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0) AS [Излишки метраж]" Else Q = Q + "IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0) AS [Излишки метраж]"
AnalizLgot.G = AnalizLgot.G + 1

End If

If Check12.Value = 1 Then
If Q <> "SELECT " Then Q = Q + ", " + "Round(IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0)*([Adding]![TarifI]+[Adding]![Tarif]),2) AS [Нач излишки]" Else Q = Q + "Round(IIf(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)>IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),Round(IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0)-IIf([Adding]![ObPl]>[Adding]![Socmin],10*[Adding]![Propis],0),1),0)*([Adding]![TarifI]+[Adding]![Tarif]),2) AS [Нач излишки]"
AnalizLgot.G = AnalizLgot.G + 1

End If


If AnalizLgot.G < 3 Then
MsgBox ("Вы не выбрали колонки отчета ")
Exit Sub
End If
'If Check.Value = 1 Then If Q <> "SELECT " Then Q = Q + ", " + "" Else Q = Q + ""


Reports.sq = Q + " FROM (Adding LEFT JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД"


If Option2.Value = True Then Reports.sq = Reports.sq + " WHERE (((IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0))<>0) AND ((Adding.KodN)=3)) OR (((Adding.KodN)=2))"
If Option1.Value = True Then Reports.sq = Reports.sq + " WHERE (((IIf([Adding]![ObPl]>[Adding]![Socmin],Round([Adding]![ObPl]-[Adding]![Socmin],1),0))<>0) AND ((Adding.KodN)=3))"

'AnalizLgot.FG1.Cols = 20
MsgBox (Reports.sq)
Unload Me
AnalizLgot.Show
AnalizLgot.Об 3

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
MsgBox ("В стадии разработки! ДЛЯ ВСЕХ КВАРТИР ОТЧЕТ БУДЕТ СОБРАН НЕ ВЕРНО, СУММЫ ИЗЛИШЕК БУДУТ ЗАВЫШЕНЫ РОВНО В 2 РАЗА. <<<<< Будет исправлено в следующей версии программы>>>>")
'Option1.Value = True
End If
'Check1.Value = 0

End Sub
