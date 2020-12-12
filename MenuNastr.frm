VERSION 5.00
Begin VB.Form MenuNastr 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4464
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   8676
   ControlBox      =   0   'False
   Icon            =   "MenuNastr.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "MenuNastr.frx":030A
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   723
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Экспорт в банк(TXT формат)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   4332
   End
   Begin VB.CommandButton Command101 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Настройка привязки ПОЛЕЙ к реестрам банка"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "П О Ч Т А"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   4335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "л/сч без начислений"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Настройка привязки начислений к реестрам банка и соцгарантий"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Двойне записи"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Выгрузка в Социальные гарантии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Реквизиты предприятия"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Сжатие БД"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1590
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   8652
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Создание архивной копии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1095
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Импорт льгот"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2085
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ввод потоком"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Экспорт в банк(DBF формат)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   4335
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
      Picture         =   "MenuNastr.frx":0614
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   156
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   960
      Picture         =   "MenuNastr.frx":085E
      Top             =   0
      Width           =   228
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
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
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   3690
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "MenuNastr.frx":0FA8
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   480
      Picture         =   "MenuNastr.frx":16F2
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "MenuNastr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrNameB As String
Public K_Imp As String


Private Sub BtnEnh1_1_Click()

End Sub

Private Sub BtnEnh1_Click()
Dim SocGR As ADODB.Recordset ' рекордсет для выгрузки начислений в Соцгарантии
Dim rsNumSG As ADODB.Recordset ' Код клиента из setting
Dim paySG As ADODB.Recordset ' Таблица-шаблон Access
Dim RsDBFsg As ADODB.Recordset ' рекордсет для DBF
Dim Kod, StrNameSG As String
Dim SG As ADODB.Recordset 'Рекордсет файла с номерами соцгарантий
Dim Bn1 As String

'Дополняем нолями номера лиц счетов
Set SG = New ADODB.Recordset
SG.Open ("SELECT SGNUM.newnum FROM SGNUM"), Mconn, adOpenKeyset, adLockPessimistic

SG.MoveFirst
  
              Do While Not SG.EOF
Bn1 = SG("NewNum")
Do While Len(Bn1) < 12
Bn1 = "0" + Bn1
Loop
SG("NewNum") = Bn1
SG.Update
                   SG.MoveNext
                   Loop
'*****************

SG.Close


' Получаем код клиента из setting
Set rsNumSG = New ADODB.Recordset
rsNumSG.Open ("SELECT Settings.TekData, Settings.Ray, Settings.Jak FROM Settings"), Mconn, adOpenKeyset, adLockPessimistic
Kod = rsNumSG("ray") + rsNumSG("Jak") + "_" + Right(Str(rsNumSG("TekData")), 7)



КоннектDBF


'Открываем рекордсет для выгрузки начислений в Соцгарантии
Set SocGR = New ADODB.Recordset
'SocGR.Open ("SELECT Kategor.kp_kind, 0 AS ls_num, [MainOccupant]![BanKN] AS synonym, [Adding]![SummaI] AS summa, 0 AS sum_dop, 0 AS sum_peni, [Adding]![DataR] AS pay_date, 0 AS dt_start, 0 AS dt_end, 0 AS barcode, [MainOccupant]![FAM]+' '+[MainOccupant]![IM]+' '+[MainOccupant]![OT] AS fio, [kls_podr]![NAIM_KLS]+' кв.№ '+[MainOccupant]![kv_num] AS address, 0 AS plpor_num, 0 AS plpor_date FROM ((Adding INNER JOIN Kategor ON Adding.KodKat = Kategor.Код) INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN kls_podr ON MainOccupant.Dom = kls_podr.КОД WHERE (((Adding.Tip)='-') AND (([Adding]![SummaI])<>0))"), Mconn

SocGR.Open ("SELECT pay.ID_PAY, pay.KP_KIND, pay.LS_NUM, pay.SYNONYM, pay.SUMMA, pay.SUM_DOP, pay.SUM_PENI, pay.PAY_DATE, pay.DT_START, pay.DT_END, pay.BARCODE, pay.FIO, pay.ADDRESS, pay.PLPOR_NUM, pay.PLPOR_DATE, pay.KodN FROM pay"), Mconn
' SocGR.Open ("SELECT pay.id_pay, pay.kp_kind, pay.ls_num, pay.synonym, pay.summa, pay.sum_dop, pay.sum_peni, pay.pay_date, pay.dt_start, pay.barcode, pay.dt_end, pay.fio, pay.address, pay.plpor_num, pay.plpor_date FROM pay"), Mconn

' Заполняем таблицу Access
Set paySG = New ADODB.Recordset
'ЧИСТИМ
paySG.Open ("DELETE pay.* FROM pay"), Mconn

'Добавляем
paySG.Open ("INSERT INTO pay SELECT SGpay.* FROM SGpay")


'Запрос на выбор кодов платежей для выгрузки

VibNac.Show (1)



'MsgBox (VibNac.Nabor(1))
'*******************************


'Формируем имя файла для выгрузки и копируем шаблон под новым именем

StrNameSG = Replace(Kod, ".", "_")


FileCopy App.Path + "/Dbf/pay.DBF", App.Path + "/dbf/" + StrNameSG + ".dbf"

' Заполняем DBF




Set RsDBFsg = New ADODB.Recordset




RsDBFsg.Open (StrNameSG + ".dbf"), DBFConn, adOpenKeyset, adLockBatchOptimistic

'DBFConn.Execute ("")


SocGR.MoveFirst
Do While Not SocGR.EOF

If SocGR("KodN") = Nabor(1) Or SocGR("KodN") = Nabor(2) Or SocGR("KodN") = Nabor(3) Or SocGR("KodN") = Nabor(4) Or SocGR("KodN") = Nabor(5) Then

RsDBFsg.AddNew
RsDBFsg("ID_PAY") = SocGR("ID_PAY")
RsDBFsg("KP_KIND") = SocGR("KP_KIND")
RsDBFsg("LS_NUM") = SocGR("ls_num") ' Код плательщика СОЦГАРАНТИЙ
RsDBFsg("dt_start") = rsNumSG("TekData")
RsDBFsg("dt_end") = rsNumSG("TekData")
RsDBFsg("SYNONYM") = SocGR("SYNONYM") ' 12 значный лиц.счет
RsDBFsg("SUMMA") = SocGR("SUMMA") 'сумма
RsDBFsg("SUM_DOP") = SocGR("SUM_DOP")
RsDBFsg("SUM_PENI") = SocGR("SUM_PENI")
RsDBFsg("PAY_DATE") = SocGR("PAY_DATE")
RsDBFsg("FIO") = SocGR("FIO")
RsDBFsg("ADDRESS") = SocGR("ADDRESS")
RsDBFsg.UpdateBatch
End If

SocGR.MoveNext
Loop
SocGR.Close
rsNumSG.Close
'MsgBox ("Данные выгружены")

MenuNastr.StrNameB = StrNameSG + ".DBF"
BImport.Show

End Sub

Private Sub BtnEnh2_Click()
Reports.sq = ""
Unload Reports
Analizlgot.Titl = "Двойные записи"
Analizlgot.G = 10
Analizlgot.StrSQL = "SELECT Adding.KodKv, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.kv_num AS [Кв №], KLS_PODR.NAIM_KLS AS Адрес, Adding.KodN AS [Код нач], Adding.NameN AS Начисление, Adding.SummaI, Adding.Key AS Сумма FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((Adding.KodKv) In (SELECT [KodKv] FROM [Adding] As Tmp GROUP BY [KodKv],[KodN],[KodKat],[SummaI] HAVING Count(*)>1  And [KodN] = [Adding].[KodN] And [KodKat] = [Adding].[KodKat] And [SummaI] = [Adding].[SummaI])) AND ((Adding.Tip)=" + Chr(34) + "+" + Chr(34) + ")) ORDER BY Adding.KodKv, Adding.KodN, Adding.NameN, Adding.SummaI"
Analizlgot.Об 0

Unload Me
Analizlgot.Show

End Sub

Private Sub BtnEnh3_Click()
MenuNastr.Enabled = False
BankNas.Show
End Sub

Private Sub Command1_Click()
Settings.Show
End Sub

Private Sub Command10_Click()
'Выгрузка данных в TXT Файл для банка

DomImp.Show 1
BANKtxt.Show


End Sub


Private Sub Command101_Click()
BankPOLE.Show 1

End Sub

Private Sub Command2_Click()
Unload MenuNastr
Potok.Show
Unload Me
End Sub

Private Sub Command3_Click()
ДисКоннект
MainForm.Сжать_Click
MenuNastr.Show
Коннект MainForm.strDataName
End Sub

Private Sub Command4_Click()
Unload Me
MainMenu.Enabled = True
MainMenu.Show

End Sub





Private Sub Command5_Click()
'ImpLg.Show
End Sub

Private Sub Command6_Click()

DomImp.Show 1



Dim rsAcs As ADODB.Recordset
Dim rsdbf As ADODB.Recordset
Dim rsSt As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim BNum As String
Dim Sum As Double
Dim bSum As Boolean

If MsgBox("Выгрузить данные в файл, для передачи в банк", vbYesNo) = vbNo Then Exit Sub





Set rsNum = New ADODB.Recordset
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Ray, MainOccupant.Jak, MainOccupant.BanKN FROM MainOccupant"), Mconn, adOpenKeyset, adLockPessimistic

If MsgBox("Выгрузить данные с суммами? !ВНИМАНИЕ!  Для выгрузки с суммами все лиц счета должны быть расчитаны", vbYesNo) = vbYes Then bSum = True Else bSum = False

'bSum = True


If bSum Then
' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo

 
  '  Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Sum(Adding.SummaI) AS [Sum-SummaI] From Adding WHERE (((Adding.Tip)='+')) GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK HAVING (((Adding.KodKat)=" + Me.K_Imp + "))")

' СУММА НАЧИСЛЕНО - SN
' САЛЬДО НАЧАЛЬНОЕ - SALDO
' САЛЬДО КОНЕЧНОЕ - SK

'Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN, Saldo ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Sum(Adding.SummaI) AS [Sum-SummaI], Sum(Adding.SaldoN) AS [Sum-SaldoN] From Adding WHERE (((Adding.Tip)='+')) GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK HAVING (((Adding.KodKat)=" + Me.K_Imp + "))")
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Sum(Adding.SummaI) AS [Sum-SummaI] From Adding WHERE (((Adding.Tip)='+')) GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK HAVING (((Adding.KodKat)=" + Me.K_Imp + "))")

End If



If bSum Then StrNameB = "I" + rsNum("Ray") + rsNum("Jak") + "S" + ".DBF" Else StrNameB = "I" + rsNum("Ray") + rsNum("Jak") + ".DBF"

rsNum.Close

Pod.Show
If bSum Then FileCopy App.Path + "/Dbf/BankEmpt.DBF", App.Path + "/dbf/" + StrNameB Else FileCopy App.Path + "/Dbf/BankEmpt1.DBF", App.Path + "/dbf/" + StrNameB

Pod.Label1.Caption = "Подождите идет выгрузка данных в файл"
КоннектDBF
Set rsdbf = New ADODB.Recordset
Set rsAcs = New ADODB.Recordset
Set rsSt = New ADODB.Recordset

rsSt.Open ("SELECT Settings.TekData, Settings.Ray, Settings.Jak FROM Settings"), Mconn

' Запрос для выгрузки с суммами


'If bSum Then rsAcs.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.[Imp], Saldo.Saldo, Saldo.KodKat, Saldo.Nach FROM (MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) LEFT JOIN Saldo ON MainOccupant.Numer = Saldo.KodKv WHERE (((KLS_PODR.[Imp])=True) AND ((Saldo.Saldo) Is Not Null) AND ((Saldo.KodKat)=" + Me.K_Imp + "))"), Mconn, adOpenKeyset, adLockPessimistic

If bSum Then rsAcs.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.[Imp], Saldo.KodKat, Saldo.SN, Saldo.SK FROM (MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) LEFT JOIN Saldo ON MainOccupant.Numer = Saldo.KodKv WHERE (((KLS_PODR.[Imp])=True) AND ((Saldo.KodKat)=" + Me.K_Imp + ") AND ((Saldo.SK) Is Not Null))"), Mconn, adOpenKeyset, adLockPessimistic
' Запрос для выгрузки без сумм
If bSum = False Then rsAcs.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.[Imp] FROM MainOccupant RIGHT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((KLS_PODR.[Imp])=True))"), Mconn, adOpenKeyset, adLockPessimistic



rsdbf.Open (StrNameB + ".DBF"), DBFConn, adOpenKeyset, adLockBatchOptimistic
Pod.Show 0
Pod.Refresh
Pod.Label1.Refresh

Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = rsAcs.RecordCount + 2
rw = 1





BNum = "1"
Sum = 0

Do While Not rsAcs.EOF

Sum = 0
rsdbf.AddNew
rsdbf("Jak") = rsSt("Ray") + rsSt("Jak")
rsdbf("DATE") = rsSt("TekData")
'rsdbf("LCH") = rsSt("")
rsdbf("NEWNUM") = rsAcs("BanKN")
rsdbf("OLDNUM") = rsAcs("OLDNUM")

' Для выгрузки с сумами

' СУММА НАЧИСЛЕНО - SN
' САЛЬДО НАЧАЛЬНОЕ - SALDO
' САЛЬДО КОНЕЧНОЕ - SK

If bSum Then

rsdbf("SALDO") = rsAcs("SK")

rsdbf("Sopl") = rsAcs("SN")
' V_Sum = V_Sum + rsAcs("SALDO")
rsAcs.UpdateBatch
End If


If rsAcs("Fam") = "" Then
rsAcs("Fam") = "_"
rsAcs.UpdateBatch
End If
If rsAcs("Ot") = "" Then
rsAcs("Ot") = "_"
rsAcs.UpdateBatch
End If
If rsAcs("Im") = "" Then
rsAcs("Im") = "_"
rsAcs.UpdateBatch

End If






If rsAcs("Fam") <> "" Then rsdbf("FIO") = rsAcs("Fam") + " " + Replace(rsAcs("Im"), "*", "") + " " + Replace(rsAcs("Ot"), "*", "")
rsdbf("ADR") = rsAcs("NAIM_KLS") + " Кв №" + rsAcs("kv_num")

Pod.ProgressBar1.Value = rw
rsAcs.MoveNext
rw = rw + 1
Loop
rsdbf.UpdateBatch


rsdbf.Close

Unload Jdite

BImport.Show
DBFConn.Close
End Sub

Private Sub Command7_Click()
Jdite.Show
ДисКоннект
If MainForm.Arhiv("Kvartplata.amd", False) Then
End If
Коннект MainForm.strDataName
Unload Jdite
MainForm.Show
MenuNastr.Show

End Sub





Private Sub Command8_Click()
Reports.sq = ""
Unload Reports
Analizlgot.Titl = "Л/сч без начислений. за " + MainMenu.Command13.Caption

Analizlgot.G = 7
Analizlgot.StrSQL = "SELECT MainOccupant.BanKN AS №, KLS_PODR.NAIM_KLS AS Улица, MainOccupant.kv_num AS Кв, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество FROM KLS_PODR INNER JOIN (MainOccupant LEFT JOIN AdNach ON MainOccupant.Numer = AdNach.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((AdNach.KodKv) Is Null))"
Analizlgot.Об 0

Unload Me
Analizlgot.Show
End Sub

Private Sub Command9_Click()


DomImp.Show 1



Dim rsAcs As ADODB.Recordset
Dim rsdbf As ADODB.Recordset
Dim rsSt As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim BNum As String
Dim Sum As Double
Dim bSum As Boolean

If MsgBox("Выгрузить данные в файл, для почты", vbYesNo) = vbNo Then Exit Sub


Set rsNum = New ADODB.Recordset
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Ray, MainOccupant.Jak, MainOccupant.BanKN FROM MainOccupant"), Mconn, adOpenKeyset, adLockPessimistic



'If MsgBox("Выгрузить данные с суммами?", vbYesNo) = vbYes Then bSum = True Else bSum = False

bSum = False

If bSum Then StrNameB = "I" + rsNum("Ray") + rsNum("Jak") + "P" + ".DBF" Else StrNameB = "I" + rsNum("Ray") + rsNum("Jak") + ".DBF"



Pod.Show

FileCopy App.Path + "/Dbf/Post_Emp.DBF", App.Path + "/dbf/" + StrNameB

Pod.Label1.Caption = "Подождите идет выгрузка данных в файл"
КоннектDBF
Set rsdbf = New ADODB.Recordset
Set rsAcs = New ADODB.Recordset
Set rsSt = New ADODB.Recordset

rsSt.Open ("SELECT Settings.TekData, Settings.Ray, Settings.Jak FROM Settings"), Mconn
rsAcs.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, KLS_PODR.NAIM_KLS, KLS_PODR.Num, Adding.KodKat, Adding.SaldoK, BankNastr.ExpPole FROM BankNastr RIGHT JOIN ((Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) ON BankNastr.KatCod = Adding.KodKat Where (((KLS_PODR.[Imp]) = Yes)) GROUP BY MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, KLS_PODR.NAIM_KLS, KLS_PODR.Num, Adding.KodKat, Adding.SaldoK, BankNastr.ExpPole ORDER BY MainOccupant.BanKN"), Mconn, adOpenKeyset, adLockPessimistic
rsdbf.Open (StrNameB + ".DBF"), DBFConn, adOpenKeyset, adLockBatchOptimistic
Pod.Refresh
Pod.Label1.Refresh

Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = rsAcs.RecordCount + 2
rw = 1

rsAcs.MoveFirst
BNum = "1"
Sum = 0
Do While Not rsNum.EOF

If BNum <> rsAcs("BanKN") Then

rsdbf.UpdateBatch

Sum = 0
rsdbf.AddNew
'rsdbf("Jak") = rsSt("Ray") + rsSt("Jak")
'rsdbf("DATE") = rsSt("TekData")
'rsdbf("LCH") = rsSt("")
rsdbf("LSHET") = rsNum("BanKN")
'rsdbf("OLDNUM") = rsAcs("OLDNUM")

If rsNum("Fam") = "" Then
rsNum("Fam") = "_"
rsNum.UpdateBatch
End If
If rsNum("Ot") = "" Then
rsNum("Ot") = "_"
rsNum.UpdateBatch
End If
If rsNum("Im") = "" Then
rsNum("Im") = "_"
rsNum.UpdateBatch
End If

If rsNum("Fam") <> "" Then rsdbf("FIO") = rsNum("Fam") + " " + Replace(rsNum("Im"), "*", "") + " " + Replace(rsNum("Ot"), "*", "")
rsdbf("ADDRESS") = rsNum("NAIM_KLS") + " Кв №" + rsNum("kv_num")

Pod.ProgressBar1.Value = rw

Zapis = rsAcs("ExpPole")

'Для экспорта без сумм устанавливаем проверку If bSum Then

If rsAcs("ExpPole") <> "" Then If bSum Then rsdbf(Zapis) = rsAcs("SaldoK")
Sum = Sum + rsAcs("SaldoK")
If bSum Then rsdbf("Sopl") = Sum


Else

Zapis = rsAcs("ExpPole")

'If rsAcs("BanKN") = "276038020408" Then MsgBox rsAcs("SaldoK")

If bSum Then If rsAcs("ExpPole") <> "" Then rsdbf(Zapis) = rsAcs("SaldoK")
Sum = Sum + rsAcs("SaldoK")
If bSum Then rsdbf("Sopl") = Sum

End If



BNum = rsAcs("BanKN")


rsAcs.MoveNext
rw = rw + 1

Loop

rsdbf.UpdateBatch


'For Rw = 1 To Dat.FG.Rows - 1
'rsdbf.AddNew
'rsdbf("R_LIC") = Dat.FG.TextMatrix(Rw, 2)
'rsdbf("R_Fio") = Dat.FG.TextMatrix(Rw, 3)
'rsdbf("R_Adr") = Dat.FG.TextMatrix(Rw, 8)
'rsdbf.UpdateBatch
'Pod.ProgressBar1.Value = Rw
'Next

rsdbf.Close

Unload Jdite
'a = Numer("906547", "55", "01")
'MsgBox Numer("906547", "55", "01")
'MsgBox ProverkaNumer("906548750151")

rsNum.Close
BImport.Show
DBFConn.Close



End Sub

'Public K_Imp As String

Private Sub Form_Load()
MakeWindow Me, True
lblTitle.Caption = "Настройки"
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


