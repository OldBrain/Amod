VERSION 5.00
Begin VB.Form RepParam 
   Caption         =   "Параметры отчета"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form8"
   ScaleHeight     =   2745
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "По видам льгот"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "По категориям расчета"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   3255
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   6
      Text            =   "Все"
      Top             =   1320
      Width           =   5775
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Text            =   "Все"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Text            =   "Все"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Адрес:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Категория расчета:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Подразделение:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "RepParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Категория As String
Dim Подразделение As String
Dim Адрес As String

Private Sub Combo1_Validate(Cancel As Boolean)
Подразделение = Trim(Combo1.Text)

End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
Категория = Combo2.Text
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
Адрес = Combo3.Text
End Sub

Private Sub Command1_Click()


If Combo1.Text = "Админ1" Then
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], Sum([Без льгот]-[Начислено]) AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.PloLG AS [Лиг пло], tmp_lgota.Procent AS [Процент льгот], [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom"

Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.PloLG, tmp_lgota.Procent, [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)='Квартплата') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS"


MsgBox Reports.sq
'Analizlgot.Об 2
Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
Exit Sub
'Unload RepLgota




End If





If Combo1.Text = "Админ" Then

If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True


Exit Sub
End If

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 10
'Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.PloLG) AS [Лиг площадь], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1))"

'Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.PloLG) AS [Лиг площадь], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], ([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100 AS [К возмещениюию] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, ([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.PloLG) AS [Лиг площадь], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], Sum(([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100) AS [К возмещениюию], Adding.ispr, Count(tmp_lgota.UniKOd) AS [Кол-во льгот] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"

Analizlgot.Об 2
Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
Exit Sub
'Unload RepLgota

End If

Unload RepLgota

If Подразделение <> "Все" Then
'Подразделение = "Like " + Chr(34) + "*" + Chr(34) + ""
wp = "((KLS_PODR.Подразделение) = " + Chr(34) + Подразделение + Chr(34) + ") And "
Else
wp = ""
End If
If Адрес <> "Все" Then
'Адрес = "Like " + Chr(34) + "*" + Chr(34) + ""
WA = "(([KLS_PODR]![NAIM_KLS] +" + Chr(34) + " дом № " + Chr(34) + "+ [KLS_PODR]![Num]) = " + Chr(34) + Адрес + Chr(34) + ") And "
Else
WA = ""
End If
If Категория <> "Все" Then
'Категория = "Like " + Chr(34) + "*" + Chr(34) + ""
WK = "((Adding.NameKat) = " + Chr(34) + Категория + Chr(34) + ") And "
Else
WK = ""
End If

Analizlgot.G = 13
If Option1.Value = True Then Reports.sq = "SELECT KLS_PODR.Подразделение, Adding.NameKat AS [Категория расчета], tmp_lgota!NAME_KLS AS Льгота, Str(tmp_lgota!Procent)+" + Chr(34) + " % " + Chr(34) + "+ tmp_lgota!Use AS [Способ прим], KLS_PODR!NAIM_KLS+" + Chr(34) + " дом № " + Chr(34) + "+KLS_PODR!Num AS Адрес, Adding.ObPl AS [Общая площадь], tmp_lgota.PloLG AS Лиг_площадь, Adding.Propis AS Прописано, Adding.SummaI AS Начислено, Round(Adding!SummaI/Adding!LgotaP,2) AS [Без льгот], Round(Adding!SummaI/Adding!LgotaP,2)-Adding!SummaI AS [К возмещению], Adding.LgotaP AS [Процент льготы], tmp_lgota.Prim AS [КолВо льгот] FROM (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) ON Adding.KodKv = MainOccupant.Numer Where "

If Option2.Value = True Then Reports.sq = "SELECT KLS_PODR.Подразделение, tmp_lgota!NAME_KLS AS Льгота, Adding.NameKat AS [Категория расчета],  Str(tmp_lgota!Procent)+" + Chr(34) + " % " + Chr(34) + "+ tmp_lgota!Use AS [Способ прим], KLS_PODR!NAIM_KLS+" + Chr(34) + " дом № " + Chr(34) + "+KLS_PODR!Num AS Адрес, Adding.ObPl AS [Общая площадь], tmp_lgota.PloLG AS Лиг_площадь, Adding.Propis AS Прописано, Adding.SummaI AS Начислено, Round(Adding!SummaI/Adding!LgotaP,2) AS [Без льгот], Round(Adding!SummaI/Adding!LgotaP,2)-Adding!SummaI AS [К возмещению], Adding.LgotaP AS [Процент льготы], tmp_lgota.Prim AS [КолВо льгот] FROM (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) ON Adding.KodKv = MainOccupant.Numer Where "

Reports.sq = Reports.sq + "(" + wp + WK + WA + "((Adding.LgotaP) < 1) And ((tmp_lgota.Prim) = 1) And ((Adding.Tip) =" + Chr(34) + "+" + Chr(34) + ")) ORDER BY Adding.NameKat, tmp_lgota!NAME_KLS, KLS_PODR!NAIM_KLS +" + Chr(34) + " дом № " + Chr(34) + " +KLS_PODR!Num"
Analizlgot.Об 3
Analizlgot.Show
Analizlgot.Caption = "Подразделение> " + Подразделение + "   Категория> " + Категория + "   Адрес> " + Адрес
Unload Me
End Sub

Private Sub Form_Load()
'Exit Sub

Dim cnParam As ADODB.Connection
Dim rsVrem As ADODB.Recordset

Option1.Value = True

Set cnParam = New ADODB.Connection
  
  cnParam.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.amd;Persist Security Info=True"
  cnParam.Open "data/Kvartplata.amd"
    
Set rsVrem = New ADODB.Recordset
Set rsVrem.ActiveConnection = cnParam
 
'Подразделение
Подразделение = "Все"
Категория = "Все"
Combo1.AddItem "Все"
For I = 1 To 10
Combo1.AddItem "Подр.№" + Trim(Str(I))
Next I

'Категория расчета


rsVrem.Open ("SELECT Kategor.Name_Kategor FROM Kategor")
rsVrem.MoveFirst
Combo2.AddItem "Все"
Do While Not rsVrem.EOF
Combo2.AddItem rsVrem.Fields("Name_Kategor")
rsVrem.MoveNext
Loop
rsVrem.Close

'Адрес
Адрес = "Все"

rsVrem.Open ("SELECT KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM KLS_PODR")
rsVrem.MoveFirst
Combo3.AddItem "Все"
Do While Not rsVrem.EOF
Combo3.AddItem rsVrem.Fields("NAIM_KLS") + " дом № " + rsVrem.Fields("Num")
rsVrem.MoveNext
Loop
rsVrem.Close


Set cnParam = Nothing
Set rsVrem = Nothing
End Sub

