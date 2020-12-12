VERSION 5.00
Begin VB.Form MenuNastr1 
   Caption         =   "Настройка"
   ClientHeight    =   7476
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   5184
   LinkTopic       =   "Form7"
   ScaleHeight     =   7476
   ScaleWidth      =   5184
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command21 
      Caption         =   "Проставить потерянное сальдо на начало"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2760
      TabIndex        =   20
      Top             =   1680
      Width           =   2172
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   372
      Left            =   3000
      TabIndex        =   19
      Top             =   6960
      Width           =   1812
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Добавить в Saldo_Arh"
      Enabled         =   0   'False
      Height          =   372
      Left            =   360
      TabIndex        =   18
      Top             =   3600
      Width           =   2172
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Убрать нулевые простои ТОЛЬКО ЛИФТ"
      Height          =   372
      Left            =   2760
      TabIndex        =   17
      Top             =   1200
      Width           =   2172
   End
   Begin VB.CommandButton Command17 
      Caption         =   "NewNum = OldNum"
      Height          =   372
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   2172
   End
   Begin VB.CommandButton Command16 
      Caption         =   "УБРАТЬ ЛИШНИЕ 0 -СТРОКИ"
      Height          =   372
      Left            =   2760
      TabIndex        =   15
      Top             =   720
      Width           =   2172
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Проверка сальдо"
      Height          =   372
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   2172
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Добавить в Saldo_arh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2760
      TabIndex        =   13
      Top             =   6000
      Width           =   2172
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Прономеровать л/сч банка"
      Height          =   735
      Left            =   2760
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Проставить сальдо на начало для записей из документов"
      Height          =   735
      Left            =   2760
      TabIndex        =   11
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Расчитать сальдо на кодец периода"
      Height          =   735
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Расчитать количество для правильного расчета сальдо"
      Height          =   735
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Проверить и добавить льготы"
      Height          =   372
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Пересчитать льготы (потребуется время)"
      Height          =   492
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Пересчет льгот таблица Tmp_Lgota, проставляет в Adding проценты наилучших льгот для быстрого расчета в дальнейшем"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Создание архивной копии"
      Height          =   372
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Выполнить после сбоев при ошибках в  расчетах"
      Height          =   855
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "Проставить ВСЕМ типы домов и квартир, прочие данные(площадь,прописано и.т.д.) из MainOccupant в Adding, а так же проставить тарифы"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Проставить постоянные нач.всем"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Выход"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Сжатие БД"
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Импорт данных из <Infin>"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Реквизиты предприятия"
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "MenuNastr1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Settings.Show
End Sub

Private Sub Command10_Click()
Jdite.Show
Jdite.Label1.Refresh
MainForm.КоличествоСальдо All
Unload Jdite
End Sub

Private Sub Command11_Click()
Jdite.Show
Jdite.Label1.Refresh
MainForm.RSaldoK "All"
Unload Jdite
End Sub

Private Sub Command12_Click()
MainForm.RSaldoN All
MsgBox ("Ok")
End Sub

Private Sub Command13_Click()
Dim rsNum As ADODB.Recordset
Pod.Show
Pod.ProgressBar1.min = 1
Коннект MainForm.strDataName
Set rsNum = New ADODB.Recordset
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Ray, MainOccupant.Jak, MainOccupant.BanKN FROM MainOccupant"), Mconn, adOpenKeyset, adLockPessimistic
Pod.Refresh


rsNum.MoveFirst
Pod.ProgressBar1.Max = 2
Do While Not rsNum.EOF
Pod.ProgressBar1.Max = Pod.ProgressBar1.Max + 1
rsNum.MoveNext
Loop


rsNum.MoveFirst
Do While Not rsNum.EOF
Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
rsNum("BankN") = Numer(rsNum("Numer"), rsNum("Jak"), rsNum("Ray"))
rsNum.UpdateBatch
rsNum.MoveNext
Loop
rsNum.Close
ДисКоннект
Unload Pod
End Sub

Private Sub Command14_Click()
Dim rsAdSaldo As ADODB.Recordset

Set rsAdSaldo = New ADODB.Recordset

rsAdSaldo.Open ("SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoN FROM Adding LEFT JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV) WHERE (((Saldo_Arh.KodKV) Is Null))"), Mconn

If rsAdSaldo.EOF = False Or rsAdSaldo.BOF = False Then
rsAdSaldo.MoveFirst

Do While Not rsAdSaldo.EOF

'MsgBox "INSERT INTO Saldo_Arh ( KodKV, KodKat, SK ) SELECT " + Str(rsAdSaldo("KodKv")) + ", " + Str(rsAdSaldo("KodKat")) + ", " + Str(rsAdSaldo("SaldoN"))

Mconn.Execute ("INSERT INTO Saldo_Arh ( KodKV, KodKat, SK ) SELECT " + Str(rsAdSaldo("KodKv")) + ", " + Str(rsAdSaldo("KodKat")) + ", " + Str(rsAdSaldo("SaldoN")))

rsAdSaldo.MoveNext
Loop
End If
End Sub

Private Sub Command15_Click()


Arhiv_all.Show
    
End Sub

Private Sub Command16_Click()
'Убираем двойные записи
Jdite.Show
' Расчитаем количество на всякий случай
MainForm.КоличествоСальдо All
Jdite.Label1.Caption = "Подождите идет расчет количества строк "

' Создаем временную таблицу
Mconn.Execute ("DELETE DUB_DEL.* FROM DUB_DEL")
Mconn.Execute ("INSERT INTO DUB_DEL ( KodKv, KodKat, SaldoN, SaldoK, Tip, KodDoc, SummaI, Kol, NameN, [Count-KodKv], Выражение1 ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.KodDoc, Adding.SummaI, Adding.Kol, Adding.NameN, Count(Adding.KodKv) AS [Count-KodKv], [Kol]-[Count-KodKv] AS Выражение1 From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.KodDoc, Adding.SummaI, Adding.Kol, Adding.NameN HAVING (((Adding.Tip)='-') AND ((Adding.KodDoc)=0) AND ((Adding.SummaI)=0) AND ((Adding.Kol)>1)) ORDER BY Adding.KodKv, Adding.KodKat")

'Mconn.Execute ("INSERT INTO DUB_DEL ( KodKv, KodKat, SaldoN, SaldoK, Tip, KodDoc, SummaI, Kol, NameN, [Count-KodKv], Выражение1 ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.KodDoc, Adding.SummaI, Adding.Kol, Adding.NameN, Count(Adding.KodKv) AS [Count-KodKv], [Kol]-[Count-KodKv] AS Выражение1 From Adding Where (((Adding.KodDoc) = 0)) GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.SummaI, Adding.Kol, Adding.NameN HAVING (((Adding.Tip)='-') AND ((Adding.SummaI)=0) AND ((Adding.Kol)>1)) ORDER BY Adding.KodKv, Adding.KodKat")


' Убираем строки из временной таблицы что бы случайно не удалить сальдо
Mconn.Execute ("DELETE DUB_DEL.*, DUB_DEL.Выражение1 From DUB_DEL WHERE (((DUB_DEL.Выражение1)=0))")
' Теперь в таблице DUB_DEL собраны ном Л.Сч. и категории которые можно смело удалять
' Открываем цыкл по этой таблице
Dim rsDubDel As ADODB.Recordset
Set rsDubDel = New ADODB.Recordset



rsDubDel.Open ("SELECT DUB_DEL.KodKv, DUB_DEL.KodKat FROM DUB_DEL"), Mconn


If rsDubDel.EOF = False Or rsDubDel.BOF = False Then

rsDubDel.MoveFirst
Do While Not rsDubDel.EOF

Jdite.Label1.Caption = "Обрабатываю л.сч. № " + Str(rsDubDel("KodKv"))
Jdite.Label1.Refresh
' Теперь непосредственно удаление
Mconn.Execute ("DELETE Adding.KodKv, Adding.KodKat, Adding.Tip, Adding.SummaI, Adding.KodDoc From Adding WHERE (((Adding.KodKv)=" + Str(rsDubDel("KodKv")) + ") AND ((Adding.KodKat)=" + Str(rsDubDel("KodKat")) + ") AND ((Adding.Tip)='-') AND ((Adding.SummaI)=0) AND ((Adding.KodDoc)=0))")
rsDubDel.MoveNext
Loop

End If
' Расчитаем количество так как записи удалены
MainForm.КоличествоСальдо All
Jdite.Label1.Caption = "Подождите идет расчет количества строк "


Unload Jdite

'*****************************************************
' Теперь табличные документы ОДН

'Убираем двойные записи
Jdite.Show
' Расчитаем количество на всякий случай
MainForm.КоличествоСальдо All
Jdite.Label1.Caption = "Подождите идет расчет количества строк "

' Создаем временную таблицу
Mconn.Execute ("INSERT INTO DUB_DEL ( KodKv, KodKat, SaldoN, SaldoK, KodDoc, SummaI, Kol, NameN, [Count-KodKv], Выражение1 ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.TablDoc, Adding.SummaI, Adding.Kol, Adding.NameN, Count(Adding.KodKv) AS [Count-KodKv], [Kol]-[Count-KodKv] AS Выражение1 From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.TablDoc, Adding.SummaI, Adding.Kol, Adding.NameN Having (((Adding.TablDoc) <> 0) And ((Adding.SummaI) = 0) And ((Adding.Kol) > 1)) ORDER BY Adding.KodKv, Adding.KodKat")

'Mconn.Execute ("INSERT INTO DUB_DEL ( KodKv, KodKat, SaldoN, SaldoK, KodDoc, SummaI, Kol, NameN, [Count-KodKv], Выражение1 ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.TablDoc, Adding.SummaI, Adding.Kol, Adding.NameN, Count(Adding.KodKv) AS [Count-KodKv], [Kol]-[Count-KodKv] AS Выражение1 From Adding Where (((Adding.TablDoc) <> 0)) GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoN, Adding.SaldoK, Adding.SummaI, Adding.Kol, Adding.NameN Having (((Adding.SummaI) = 0) And ((Adding.Kol) > 1)) ORDER BY Adding.KodKv, Adding.KodKat")


' Убираем строки из временной таблицы что бы случайно не удалить сальдо
Mconn.Execute ("DELETE DUB_DEL.*, DUB_DEL.Выражение1 From DUB_DEL WHERE (((DUB_DEL.Выражение1)=0))")
' Теперь в таблице DUB_DEL собраны ном Л.Сч. и категории которые можно смело удалять
' Открываем цыкл по этой таблице

Set rsDubDel = New ADODB.Recordset
rsDubDel.Open ("SELECT DUB_DEL.KodKv, DUB_DEL.KodKat FROM DUB_DEL"), Mconn

If rsDubDel.EOF = False Or rsDubDel.BOF = False Then

rsDubDel.MoveFirst
Do While Not rsDubDel.EOF

Jdite.Label1.Caption = "/ОДН/ Обрабатываю л.сч. № " + Str(rsDubDel("KodKv"))
Jdite.Label1.Refresh
' Теперь непосредственно удаление
'Mconn.Execute ("DELETE Adding.KodKv, Adding.KodKat, Adding.Tip, Adding.SummaI, Adding.KodDoc From Adding WHERE (((Adding.KodKv)=" + Str(rsDubDel("KodKv")) + ") AND ((Adding.KodKat)=" + Str(rsDubDel("KodKat")) + ") AND ((Adding.Tip)='-') AND ((Adding.SummaI)=0) AND ((Adding.KodDoc)=0))")

Mconn.Execute ("DELETE Adding.KodKv, Adding.KodKat, Adding.Tip, Adding.SummaI, Adding.TablDoc From Adding WHERE (((Adding.KodKv)=" + Str(rsDubDel("KodKv")) + ") AND ((Adding.KodKat)=" + Str(rsDubDel("KodKat")) + ") AND ((Adding.Tip)='+') AND ((Adding.SummaI)=0) AND ((Adding.TablDoc)<>0))")

rsDubDel.MoveNext
Loop

End If

'Теперь проверяем потерянное сальдо ))

Set rsTheLostSaldo = New ADODB.Recordset
rsTheLostSaldo.Open ("SELECT Saldo_Arh.KodKV as KodKV, Saldo_Arh.KodKat as KodKat, Saldo_Arh.SK as SaldoN FROM Saldo_Arh LEFT JOIN ADDING ON Saldo_Arh.KodKV = ADDING.KodKv WHERE (((ADDING.KodKv) Is Null))"), Mconn

'Добавляем стороки потерянного сальдо

If rsTheLostSaldo.EOF = False Or rsTheLostSaldo.BOF = False Then 'Проверка что рекорсет не пустой

Jdite.Label1.Caption = ("Похоже что то потеряли. Сейчас восстановим!")
Jdite.Label1.Refresh

Set rsAdding = New ADODB.Recordset

rsAdding.Open ("SELECT ADDING.* FROM ADDING"), Mconn, adOpenDynamic, adLockPessimistic

rsTheLostSaldo.MoveFirst
Do While Not rsTheLostSaldo.EOF
rsAdding.AddNew
rsAdding.Fields("KodKV") = rsTheLostSaldo("KodKV")
rsAdding.Fields("KodKat") = rsTheLostSaldo("KodKat")
rsAdding.Fields("SaldoN") = rsTheLostSaldo("SaldoN")
rsAdding.Fields("NameN") = "Скоректированное сальдо"
rsAdding.Fields("SummaI") = 0
rsAdding.Fields("SaldoK") = rsTheLostSaldo("SaldoN")
rsAdding.Fields("SaldoN") = rsTheLostSaldo("SaldoN")
rsAdding.Fields("KodN") = "100" + rsTheLostSaldo("KodKat")
rsAdding.Fields("tip") = "-"
rsAdding.Fields("Formula") = "SummaI"


rsAdding.Fields("KodDoc") = 0
rsAdding.Fields("Sch") = "Нет"







rsAdding.Fields("SchetZ") = "Общие"
rsAdding.Update
rsTheLostSaldo.MoveNext
Loop
rsAdding.UpdateBatch
rsAdding.Close
End If

' Заполняем пустые значения после добавления строк для корректировки сальдо
Mconn.Execute ("UPDATE ADDING INNER JOIN Kategor ON ADDING.KodKat = Kategor.Код SET ADDING.NameKat = [Kategor]![Name_Kategor]")
Mconn.Execute ("UPDATE ADDING INNER JOIN nachisleniy ON ADDING.KodN = nachisleniy.Kod SET ADDING.NameN = [nachisleniy]![Naim], ADDING.Formula = [nachisleniy]![Formula], ADDING.Tip = [nachisleniy]![Tip]")

' Расчитаем количество так как записи удалены
MainForm.КоличествоСальдо All
Jdite.Label1.Caption = "/ОДН/ Подождите идет расчет количества строк "


Unload Jdite


MsgBox ("ОК")



End Sub

Private Sub Command17_Click()
Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.BanKN = [MainOccupant]![OLDNUM]")
End Sub

Private Sub Command18_Click()
'Убираем лишние записи


If MsgBox("ВНИМАНИЕ!! Будут Удаленв ВСЕ НУЛЕВЫЕ начисления с кодом 10", vbYesNo) = vbYes Then
Jdite.Show
Mconn.Execute ("DELETE Adding.KodN, Adding.SummaBl, Adding.SummaI From Adding WHERE (((Adding.KodN)=10) AND ((Adding.SummaBl)=0) AND ((Adding.SummaI)=0))")
'Mconn.Execute ("INSERT INTO Saldo_Arh ( KodKV, KodKat, SK ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoN From Adding WHERE (((Adding.KodKv)=" + Str(kvit("KodKv")) + ") AND ((Adding.KodKat)=" + Str(kvit("KodKat")) + "))")
Unload Jdite
End If
End Sub

Private Sub Command2_Click()
'menunastr.hide
Form1.Show

End Sub

Private Sub Command20_Click()
Form5.Show
End Sub

Private Sub Command21_Click()
'Теперь проверяем потерянное сальдо ))

Set rsTheLostSaldo = New ADODB.Recordset
rsTheLostSaldo.Open ("SELECT Saldo_Arh.KodKV as KodKV, Saldo_Arh.KodKat as KodKat, Saldo_Arh.SK as SaldoN FROM Saldo_Arh LEFT JOIN ADDING ON Saldo_Arh.KodKV = ADDING.KodKv WHERE (((ADDING.KodKv) Is Null))"), Mconn

'Добавляем стороки потерянного сальдо

If rsTheLostSaldo.EOF = False Or rsTheLostSaldo.BOF = False Then 'Проверка что рекорсет не пустой

Jdite.Label1.Caption = ("Похоже что то потеряли. Сейчас восстановим!")
Jdite.Label1.Refresh

Set rsAdding = New ADODB.Recordset

rsAdding.Open ("SELECT ADDING.* FROM ADDING"), Mconn, adOpenDynamic, adLockPessimistic

rsTheLostSaldo.MoveFirst
Do While Not rsTheLostSaldo.EOF
rsAdding.AddNew
rsAdding.Fields("KodKV") = rsTheLostSaldo("KodKV")
rsAdding.Fields("KodKat") = rsTheLostSaldo("KodKat")
rsAdding.Fields("SaldoN") = rsTheLostSaldo("SaldoN")
rsAdding.Fields("NameN") = "Скоректированное сальдо"
rsAdding.Fields("SummaI") = 0
rsAdding.Fields("SaldoK") = rsTheLostSaldo("SaldoN")
rsAdding.Fields("SaldoN") = rsTheLostSaldo("SaldoN")
rsAdding.Fields("KodN") = "100" + rsTheLostSaldo("KodKat")
'MsgBox ("10" + rsTheLostSaldo("KodKat"))
rsAdding.Fields("tip") = "-"
rsAdding.Fields("Formula") = "SummaI"
rsAdding.Fields("KodDoc") = 0
rsAdding.Fields("Sch") = "Нет"
rsAdding.Fields("SchetZ") = "Общие"
rsAdding.Update
rsTheLostSaldo.MoveNext
Loop
rsAdding.UpdateBatch
rsAdding.Close
End If

' Заполняем пустые значения после добавления строк для корректировки сальдо
Mconn.Execute ("UPDATE ADDING INNER JOIN Kategor ON ADDING.KodKat = Kategor.Код SET ADDING.NameKat = [Kategor]![Name_Kategor]")
Mconn.Execute ("UPDATE ADDING INNER JOIN nachisleniy ON ADDING.KodN = nachisleniy.Kod SET ADDING.NameN = [nachisleniy]![Naim], ADDING.Formula = [nachisleniy]![Formula], ADDING.Tip = [nachisleniy]![Tip]")

MsgBox ("ОК")

End Sub

Private Sub Command3_Click()
'conn.Close
'SetConn.Close
'Unload MainForm
'If Dir("f:\kv\amod\data\kvartplata.ldb") <> "" Then MsgBox ("В сети") Else MsgBox ("В сети ни кого нет")
'MainForm.Show

'MsgBox (App.Path)

'Set conn = Nothing
'Set SetConn = Nothing

'If gflngCompactDatabase(App.Path + "\data\kvartplata.mdb", True) Then MsgBox ("Ok") Else MsgBox ("Bad")
'MainForm.Enabled = True
MainForm.Сжать_Click
Load MainForm
MenuNastr.Show
End Sub

Private Sub Command4_Click()
Unload Me
MainMenu.Enabled = True
MainMenu.Show
End Sub

Private Sub Command5_Click()

MainForm.AddConstanta

End Sub

Private Sub Command6_Click()
MainForm.AddingTIP
End Sub

Private Sub Command7_Click()
Jdite.Show
If MainForm.Arhiv("Kvartplata.mdb", True) Then
End If
Unload Jdite
MainForm.Show
MenuNastr.Show

End Sub

Private Sub Command8_Click()
Form7.Show
End Sub

Private Sub Command9_Click()
DobLgot.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True
MainMenu.Show
End Sub

'ФУНКЦИЯ СЖАТИЯ БД DAO-Методом'
'  gflngCompactDatabase(...)'
'ВХОДНЫЕ ПАРАМЕТРЫ ФУНКЦИИ:'
'  CompactingDBPathAndName - строковый параметр, задающий ПОЛНЫЙ ПУТЬ (путь + имя файла)'
'     к сжимаемой БД.'
'  BackupBeforeCompactDB - необязательный логический параметр, указывающий на'
'     необходимость сделать перед сжатием резервную копию сжимаемой БД (резервная'
'     копия выкладывается в файл с именем "ИмяСжимаемогоФайла_Backup"). При'
'     отсутствии параметра резервное копирование не производится.'
'ВОЗВРАЩАЕМОЕ ФУНКЦИЕЙ ЗНАЧЕНИЕ:'
'  = 0, если сжатие произведено;'
'  = Номеру возникшей ошибки, если выполнить сжатие не удалось.'
'ОСОБЕННОСТИ:'
'  Для выполнения процедуры сжатия автоматически создается временный файл'
'     с именем "ПолныйПуть\ИмяСжимаемогоФайла_Temp".'
'  Резервное копирование, выполнение которого определяется параметром "BackupBeforeCompactDB",'
'     производится в файл с именем "ПолныйПуть\ИмяСжимаемогоФайла_Backup"), при'
'     этом старая копия резерва перезаписывается новой (фактически удаляется).'
'  В случае, если сжимаемая БД открыта, то файл БД не будет скопирован (соответствующая'
'     ошибка появится в момент копирования БД).'
Public Function gflngCompactDatabase( _
CompactingDBPathAndName As String, _
Optional BackupBeforeCompactDB As Boolean = False) As Long
Dim strTempFile As String

'MsgBox ("Ok+Ok")

'On Error GoTo ErrHandler
'Формируем имя для временного ("принимающего") файла'
  strTempFile = Left(CompactingDBPathAndName, (Len(CompactingDBPathAndName) - 4)) & _
  "_Temp" & Right(CompactingDBPathAndName, 4)
'Создаем (если надо) резервную копию файла БД перед сжатием'
  If BackupBeforeCompactDB = True _
  Then FileCopy CompactingDBPathAndName, _
  Left(CompactingDBPathAndName, (Len(CompactingDBPathAndName) - 4)) & _
  "_Backup" & Right(CompactingDBPathAndName, 4)
'Сжимаем файл БД (с перезаписью сжатого файла в новый файл)'
  DBEngine.CompactDatabase CompactingDBPathAndName, strTempFile, dbLangCyrillic
'Перезаписываем сжатый (временный файл) на место несжатого (старого файла)'
  FileCopy strTempFile, CompactingDBPathAndName
'Удаляем временный файл'
  Kill strTempFile
Exit Function
ErrHandler:
'обрабатываем возможные ошибки'
  gflngCompactDatabase = Err.Number
  MsgBox (Err.Description)
  Err.Clear: Exit Function
End Function


