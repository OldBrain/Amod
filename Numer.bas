Attribute VB_Name = "ConnModule"
Option Explicit
Public Mconn As ADODB.Connection
Public Zconn As ADODB.Connection
Public Ca As ADODB.Connection
Public DBFConn As ADODB.Connection
Public BestConn As ADODB.Connection
Public rsA As ADODB.Recordset
Public Arhiv As Boolean
Public result As Boolean
Public Prostoy As Boolean
Public Fn As String
Public Nabor(5) As String




Public Sub Коннект(strSur As String)
Set Mconn = New ADODB.Connection
Call BaseUnProtect(App.Path + "\data\" + strSur, True)
Mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/" + strSur + ";Jet OLEDB:Database Password=" + MainForm.Pas + ";"
Mconn.Open
'Call BaseProtect(App.Path + "\data\" + strSur, True)


End Sub
Public Sub КоннектЗ()
Set Zconn = New ADODB.Connection
Call BaseUnProtect(App.Path + "\data\zatrat.mdb", True)
Zconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/" + "zatrat.mdb" + ";Jet OLEDB:Database Password=" + MainForm.Pas + ";"
Zconn.Open
'Call BaseProtect(App.Path + "\data\" + strSur, True)


End Sub

Public Sub КоннектDBF()
Set DBFConn = New ADODB.Connection

  DBFConn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы dBASE;Initial Catalog=" + App.Path + "/dbf/"
  
 'DBFConn.Open "BASE_GH.DBF"
  End Sub
'Public Sub ДисКоннект()
'Conn.Close
'Set Conn = Nothing
'End Sub
Public Sub Best()
Set BestConn = New ADODB.Connection


  'BestConn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы dBASE;Initial Catalog=" + Trim(MainForm.BestPath)
  BestConn.Open "Provider=VFPOLEDB.1;Data Source=" + Trim(MainForm.BestPath) + ";Password='';Collating Sequence=RUSSIAN"

 'BestConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Trim(MainForm.BestPath) + ";Extended Properties=dBASE IV;User ID=Admin;Password=;"
 'BestConn.Open "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;" + Trim(MainForm.BestPath) + ";"
 
  End Sub

Public Sub ДисКоннект()
 
On Error GoTo e
Mconn.Close
Set Mconn = Nothing
e:
End Sub
Public Sub ДисКоннектЗ()
 
On Error GoTo e
Zconn.Close
Set Zconn = Nothing
e:
End Sub



       
   ' Переключить на русский язык
   'X = ActivateKeyboardLayout&(kb_lay_ru, 0)

   ' Переключить на английский язык
   'X = ActivateKeyboardLayout&(kb_lay_en, 0)


'ФУНКЦИЯ ВЫЧИСЛЕНИЯ КОНТРОЛЬНОЙ СУММЫ ЧМСЛА ДЛЯ КОДИРОВАНИЯ Л/СЧ

'Номер лицевого счета. Распишу требования к проверке, из них понятно как производится ключевание счета.

'Шаблон  - LLLLLLGGGGKK *, где

'LLLLLL  - номер Абонента;
'GGGG        - номер ЖЭКа
'               (Например: 0404 -ЖЭК №4 Трусовского р-на
'                          0501 -ЖЭК №5 Кировского р-на);
'KK      - контрольные числа.



'1.  Если введено менее трех цифр - запрет ввода значения.
'2.  Введенное значение номера л/сч. дополняется слева нулями до 12-ти символов.
'3.  Первые 10 цифр номера счета поразрядно складываются с цифрами весового ряда {3,5,7,9,3,5,7,9,3,5}, полученные суммы складываются. Последняя цифра полученного числа является ВТОРОЙ цифрой контрольного числа.
'4.  Первые 10 цифр номера счета поразрядно умножаются на цифры весового ряда {3,5,7,9,3,5,7,9,3,5}, полученные произведения складываются. Последняя цифра полученного числа является ПЕРВОЙ цифрой контрольного числа.
'5.  Происходит сравнение первой контрольной цифры с 11-ой цифрой номера лицевого счета и второй контрольной цифры с 12-ой цифрой номера лицевого счета. Далее два варианта:
'-   сравнение прошло. Имеющиеся нули с левой стороны заменяются на пробелы. Конец контроля;
'-   сравнение не прошло. Запрет ввода.


'Примеры  лицевых счетов для контроля алгоритма.
'906547550168, 906549950144, 906548750151.


'*номер Абонента лучше поставить впереди, тогда не нужно будет вводить все 12 цифр, т.к. при проверке л/сч. номер с левой стороны все равно будет добавляться нулями.

'Num-это номер Абонента внутри жэка
'Jak- Номер ЖЭКа "ХХ" два символа
'Ray-Код района   "ХХ" два символа

'Значение которое возвращает функция и есть номер л/сч по Шаблону  - LLLLLLGGGGKK

Public Function Numer(ByVal Num As String, Jak As String, Ray As String) As String
Dim VesR(10) As Integer
Dim n(10) As Double
Dim i As Integer
Dim StNlic As String
Dim Summ As Double
Dim Pro As Double
Dim First As String
Dim Too As String


 VesR(1) = 3
 VesR(2) = 5
 VesR(3) = 7
 VesR(4) = 9
 VesR(5) = 3
 VesR(6) = 5
 VesR(7) = 7
 VesR(8) = 9
 VesR(9) = 3
 VesR(10) = 5

StNlic = Trim(Num) + Trim(Jak) + Trim(Ray)

 n(1) = Int(Val(StNlic) / 1000000000)
 n(2) = Int(Val(StNlic) / 100000000) - n(1) * 10
 n(3) = Int(Val(StNlic) / 10000000) - n(1) * 100 - n(2) * 10
 n(4) = Int(Val(StNlic) / 1000000) - n(1) * 1000 - n(2) * 100 - n(3) * 10
 n(5) = Int(Val(StNlic) / 100000) - n(1) * 10000 - n(2) * 1000 - n(3) * 100 - n(4) * 10
 n(6) = Int(Val(StNlic) / 10000) - n(1) * 100000 - n(2) * 10000 - n(3) * 1000 - n(4) * 100 - n(5) * 10
 n(7) = Int(Val(StNlic) / 1000) - n(1) * 1000000 - n(2) * 100000 - n(3) * 10000 - n(4) * 1000 - n(5) * 100 - n(6) * 10
 n(8) = Int(Val(StNlic) / 100) - n(1) * 10000000 - n(2) * 1000000 - n(3) * 100000 - n(4) * 10000 - n(5) * 1000 - n(6) * 100 - n(7) * 10
 n(9) = Int(Val(StNlic) / 10) - n(1) * 100000000 - n(2) * 10000000 - n(3) * 1000000 - n(4) * 100000 - n(5) * 10000 - n(6) * 1000 - n(7) * 100 - n(8) * 10
 n(10) = Val(Right(StNlic, 1))
 
Numer = ""
Summ = 0
Pro = 0

For i = 1 To 10
Summ = Summ + n(i) + VesR(i)
Pro = Pro + n(i) * VesR(i)
Numer = Numer + Trim(Str(n(i)))
Next

Too = Right(Str(Int(Summ)), 1)
First = Right(Str(Int(Pro)), 1)
Numer = Numer + First + Too

End Function
' Функция проверки номера лиц счета (переменная Num), закодированного Функцией Numer()
'на корректность по вышеприведенному алгоритму
' расчета контрольной суммы ProverkaNumer=True Если номер верный, и False если нет

Public Function ProverkaNumer(ByVal Num As String) As Boolean
Dim VesR(10) As Integer
Dim n(10) As Double
Dim i As Integer
Dim StNlic As String
Dim Summ As Double
Dim Pro As Double
Dim First As String
Dim Too As String


 VesR(1) = 3
 VesR(2) = 5
 VesR(3) = 7
 VesR(4) = 9
 VesR(5) = 3
 VesR(6) = 5
 VesR(7) = 7
 VesR(8) = 9
 VesR(9) = 3
 VesR(10) = 5

Do While Len(Num) < 12
Num = "0" + Num
Loop

'MsgBox Num

StNlic = Left(Num, 10)


 n(1) = Int(Val(StNlic) / 1000000000)
 n(2) = Int(Val(StNlic) / 100000000) - n(1) * 10
 n(3) = Int(Val(StNlic) / 10000000) - n(1) * 100 - n(2) * 10
 n(4) = Int(Val(StNlic) / 1000000) - n(1) * 1000 - n(2) * 100 - n(3) * 10
 n(5) = Int(Val(StNlic) / 100000) - n(1) * 10000 - n(2) * 1000 - n(3) * 100 - n(4) * 10
 n(6) = Int(Val(StNlic) / 10000) - n(1) * 100000 - n(2) * 10000 - n(3) * 1000 - n(4) * 100 - n(5) * 10
 n(7) = Int(Val(StNlic) / 1000) - n(1) * 1000000 - n(2) * 100000 - n(3) * 10000 - n(4) * 1000 - n(5) * 100 - n(6) * 10
 n(8) = Int(Val(StNlic) / 100) - n(1) * 10000000 - n(2) * 1000000 - n(3) * 100000 - n(4) * 10000 - n(5) * 1000 - n(6) * 100 - n(7) * 10
 n(9) = Int(Val(StNlic) / 10) - n(1) * 100000000 - n(2) * 10000000 - n(3) * 1000000 - n(4) * 100000 - n(5) * 10000 - n(6) * 1000 - n(7) * 100 - n(8) * 10
 n(10) = Val(Right(StNlic, 1))

'Расчет контрольной суммы
'Numer = ""
Summ = 0
Pro = 0

For i = 1 To 10
Summ = Summ + n(i) + VesR(i)
Pro = Pro + n(i) * VesR(i)

Next

Too = Right(Str(Int(Summ)), 1)
First = Right(Str(Int(Pro)), 1)


If Val(Right(Num, 2)) = Int(Val(First + Too)) Then ProverkaNumer = True Else ProverkaNumer = False
End Function

Public Sub КоннектАрхив(strSur As String)

Set Ca = New ADODB.Connection


Call BaseUnProtect(App.Path + "\data\Arhiv\" + strSur, True)

Ca.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/arhiv/" + strSur + ";Jet OLEDB:Database Password=" + MainForm.Pas + ";"
On Error GoTo erc
Ca.Open

'Call BaseProtect(App.Path + "\data\Arhiv\" + strSur, True)

erc:
If Err.Number = -2147467259 Or Err.Number = 3021 Or Err.Number = 13 Or Err.Number = 0 Then
Err.Clear
Exit Sub
Else
MsgBox Err.Description + " Ошибка №" + Str(Err.Number)
Err.Clear
Exit Sub
End If


End Sub

Public Sub КоннектА(strSur As String, Номер As String, addAdding As Boolean, addingAll As Boolean)


result = False
If Номер = "" And Номер = "Номер" Then Exit Sub
Arc.n = 0
Arc.O = 0
Arc.s = 0
Arc.i = 0


Set Ca = New ADODB.Connection
Set rsA = New ADODB.Recordset

Call BaseUnProtect(App.Path + "\data\arhiv\" + strSur, True)



Ca.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/arhiv/" + strSur + ";Jet OLEDB:Database Password=" + MainForm.Pas + ";"

On Error GoTo Cl

Ca.Open



If addingAll = True Then
' Создаем временную таблицу для архивных отчетов ДЛЯ ВСЕХ
Call BaseUnProtect(App.Path + "\data\kvartplata.amd", True)
'Чистим таблицу
'Mconn.Execute ("DELETE Arh_Rep_All.* FROM Arh_Rep_All")
'Копируем ВСЕ данные
Mconn.Execute ("INSERT INTO arh_rep_all SELECT Adding.* FROM Adding IN '" + App.Path + "\Data\arhiv\" + strSur + "'")
 End If



If addAdding = True Then
' Создаем временную таблицу для архивных отчетов предварительно сняв защиту
Call BaseUnProtect(App.Path + "\data\kvartplata.amd", True)
'добавляем данные архива
Mconn.Execute ("INSERT INTO arh_rep SELECT Adding.* FROM Adding IN '" + App.Path + "\Data\arhiv\" + strSur + "' WHERE (((Adding.KodKv)=" + Номер + "));")

 End If




rsA.Open ("SELECT Adding.KodKv, round(Sum([Adding]![SaldoN]/[Adding]![Kol]),2) AS СальдоН, Sum(IIf([Adding]![Tip]='+',[Adding]![SummaI],0)) AS Начислено, Sum(IIf([Adding]![Tip]='-',[Adding]![SummaI],0)) AS Оплачено, Sum(IIf([Adding]![Tip]='s',[Adding]![SummaI],0)) AS Субсидии, round(Sum([Adding]![SaldoK]/[Adding]![Kol]),2) AS Сальдо From Adding GROUP BY Adding.KodKv HAVING (((Adding.KodKv)=" + Номер + "))"), Ca



'If rsA.RecordCount <> 0 Then
Arc.I1 = rsA("СальдоН")
Arc.n = rsA("Начислено")
Arc.O = rsA("Оплачено")
Arc.s = rsA("Субсидии")
Arc.i = rsA("Сальдо")
'End If

Cl:
If Err.Number = -2147467259 Or Err.Number = 3021 Or Err.Number = 13 Or Err.Number = 0 Then
Err.Clear
Exit Sub
Else
MsgBox Err.Description + " Ошибка №" + Str(Err.Number)
Err.Clear

result = True
Exit Sub
End If



 
rsA.Close
Ca.Close


End Sub
Public Function CheckNull(sCheck, default As String) As String
'Проверка функций на IsNull
If IsNull(sCheck) Then
CheckNull = default
Else
CheckNull = Trim$(sCheck)
End If
End Function


Public Sub FSize(ByVal NameForm As Form) 'РАЗМЕР ФОРМЫ
Dim H As Double
Dim w As Double
Dim ah As Double
Dim aw As Double
Dim Zoom As Double


With NameForm
'.Caption = TI
H = .Height
w = .Width

ah = Screen.Height
aw = Screen.Width
'.Width = aw * 0.55
.Width = aw * 1
'.Height = (h * (aw / w)) * 0.55
.Height = (H * (aw / w)) * 1
'Zoom = aw / w * 100# * 0.55
Zoom = aw / w * 100# * 1
.Top = 0#
.Left = (aw - .Width) + 20#
End With

End Sub

'модификация заголовка файла БД
'Ex: Call BaseProtect("C:\01.mdb", True)

'Public Sub BaseProtect(sPath As String, bLock As Boolean)
   ' Dim iFn As Integer
   ' iFn = FreeFile()
   ' Open sPath For Binary Access Write As #iFn
   ' Put #iFn, 5, CStr(IIf(bLock, _
   ' "Stаndard Jet DB", "Standard Jet DB"))
        '"ProtectDataBase", "Standard Jet DB"))
       
    'Close #iFn
'End Sub


Public Sub BaseUnProtect(sPath As String, bLock As Boolean)
    Dim iFn As Integer
    iFn = FreeFile()
    Open sPath For Binary Access Write As #iFn
    Put #iFn, 5, CStr(IIf(bLock, _
     "Standard Jet DB", "Stаndard Jet DB"))
       ' "ProtectDataBase", "Standard Jet DB"))
       
    Close #iFn
End Sub
