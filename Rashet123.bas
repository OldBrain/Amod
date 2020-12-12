Attribute VB_Name = "Rashet"
Option Explicit
Dim ProcentLMAX As Double ' Для подсчета общего процента льгот для лиц счета /ПРОМЕЖУТОЧНЫЙ ИТОГ/
Dim KolL As Integer ' Для расчета количества льготчиков /ПРОМЕЖУТОЧНЫЙ ИТОГ/
Dim P As Double '/ПРОМЕЖУТОЧНЫЙ ИТОГ/ для %
Dim PSumm As Double '/ПРОМЕЖУТОЧНЫЙ ИТОГ/ для СУММАРНОГО %
Dim PloSumm As Double '/ПРОМЕЖУТОЧНЫЙ ИТОГ/ для подсчета суммарной лиготируемой площади
Dim Ostat As Double
Dim SumLGPl As Double
Dim strLg As String
Dim Max As Double '/ПРОМЕЖУТОЧНЫЙ ИТОГ/ для определения максимальной(лучшей) льготы
Dim I As Integer
Dim j As Integer
Dim LastIn As Integer ' Для запоменания предыдущего номера записи рекордсета ДВОЙНЫХ льгот
Dim Dbl As Double
Dim PromItog As Double
Dim DBLCode(100) As String ' Массив для занесения  TmpLG("OtheCode")
Dim DimSocmin(100) As Integer
Dim Mpovt(100) As Integer
Dim Ind(100) As Integer
Dim Itog As Double
Dim n As Double
Dim DOP As Double
'Dim LgPlo As Double
'Dim SumPlo As Double
'Dim Socmin As Double
'Dim Propis As Double
Dim rsSocmin As ADODB.Recordset ' Для Определения соцминимума на 1-го
Dim TmpLG As ADODB.Recordset ' Все льготы по данной категории расчета
'Dim rsTmpDBL As ADODB.Recordset ' Для ДВОЙНЫХ ЛЬГОТ

Public Sub Расчет(Выбор As String)

'DoEvents







PloSumm = 0

'Обноляем параметры расчета
Mconn.Execute ("UPDATE TMP_Lgota SET TMP_Lgota.Prim = 0, TMP_Lgota.PloLG = 0, TMP_Lgota.Prim1 = 0, TMP_Lgota.itog1 = 1, TMP_Lgota.SovmPloLG = 0 WHERE (((TMP_Lgota.UniKOd)=" + Выбор + "))")

Set TmpLG = New ADODB.Recordset
'Set rsTmpDBL = New ADODB.Recordset

TmpLG.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.Plo, tmp_lgota.Prop, tmp_lgota.Cocmin, tmp_lgota.OtheCode, tmp_lgota.parametr, tmp_lgota.itog, tmp_lgota.tarif, tmp_lgota.Itog1, tmp_lgota.Prim, tmp_lgota.PloLG, tmp_lgota.Key, tmp_lgota.Prim1, tmp_lgota.Koll, tmp_lgota.SovmPloLG, Adding.KodKat, Adding.dop FROM Adding RIGHT JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd WHERE (((tmp_lgota.UniKOd)=" + Выбор + "))"), Mconn, adOpenKeyset, adLockPessimistic

'Определяем соцминимум не 1-го для данной категории расчета
Set rsSocmin = New ADODB.Recordset
rsSocmin.Open ("SELECT Socmin.KodKategor, Socmin.koli, Socmin.Value From Socmin WHERE (((Socmin.koli)=1))"), Mconn

rsSocmin.MoveFirst
Do While Not rsSocmin.EOF
DimSocmin(rsSocmin("KodKategor")) = rsSocmin("Value") + TmpLG("Dop")
rsSocmin.MoveNext
Loop






' ******* Информация для выбора льгот для отчетов

' По умолчанию в TmpLG("Prim")=0 - Запись еще не обробатывалась
'                TmpLG("Prim")=-1 - Льгота не применяется
'                TmpLG("Prim")=1 - Льгота применяется



'                                   ХОД МОИХ РАССУЖДЕНИЙ

'                   ДЛЯ НАЧИСЛЕНИЙ РАСЧИТЫВАЕМЫХ ОТ КОЛИЧЕСТВА ЧЕЛОВЕК ДОСТАТОЧНО ПРОСТО

' 1. Если TmpLG("Use")="На 1-го"  или TmpLG("Use")="На всех" то ЭТО РАСЧЕТ ОТ КОЛИЧЕСТВА ПРОПИСАННЫХ
'    значит основным параметром для расчета льготы будет являться TmpLG("Prim1").
' 2. Если TmpLG("Use")="На 1-го" то Prim1=1, если TmpLG("Use")="На всех" то Prim1=Prop (кол-ву прописанных)
' 3. Далее если параметр TmpLG("OtheCode") имеет одинаковое значение то приоритет применения имеет
'    льготв с большим значением Prim1*Procent
' 4. Если льгота применяется то проставляем TmpLG("Prim")=1 иначе TmpLG("Prim")=0

'                    Подсчитываем суммарный процент льготы для лиц, счета
' 1. Если TmpLG("Prim")=0 то пропускаем
' 2. Если TmpLG("Prim")=1 то считаем ?????


' Проставляем значения Prim1(кол-во чел.на которых применились льготы)
' и значения PloLg

'If TmpLG.RecordCount = 0 Then Exit Sub
'************* Если нет записей о льготах в файле TmpLgota to ВЫХОД****
On Error GoTo en
en:
If Err.Number = 3021 Then
'MsgBox Err.Description
Err.Clear
strLg = "0"
Mconn.Execute ("UPDATE Adding SET Adding.LgotaP =1, Adding.LgotaKod = " + Chr(34) + strLg + Chr(34) + " WHERE (((Adding.Key)=" + Выбор + "))")
Exit Sub
End If
'******************************************************************

TmpLG.MoveFirst
Do While Not TmpLG.EOF

If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("PloLg") = 0
TmpLG("Koll") = 0
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")
TmpLG.UpdateBatch
TmpLG.MoveNext
If TmpLG.EOF Then Exit Do
'Exit Sub
End If


'Если нет прописанных или общей площади то выход
If TmpLG("Plo") = 0 And TmpLG("Use") <> "На 1-го" And TmpLG("Use") <> "На всех" Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
Exit Sub
End If

If TmpLG("Prop") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch

Exit Sub
End If
'MsgBox Str(TmpLG.AbsolutePosition) + " " + Выбор


'Если нет тарифа то выход
If TmpLG("Tarif") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch

Exit Sub
End If






'------------------------- На 1-го ------------------------------
              If TmpLG("Use") = "На 1-го" Then
TmpLG("Prim1") = 1
TmpLG("PloLG") = 0
PloSumm = 0
If Not TmpLG("OtheCode") Then DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")
TmpLG("Koll") = 1
TmpLG.UpdateBatch
                         End If
'-------------------------------------------------------

'------------------------- На всех ------------------------------
              If TmpLG("Use") = "На всех" Then
TmpLG("Prim1") = TmpLG("Prop")
TmpLG("PloLG") = 0
TmpLG("Koll") = TmpLG("Prop")

TmpLG.UpdateBatch
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

                         End If
'-------------------------------------------------------

'------------------------- Вся площадь ------------------------------
              If TmpLG("Use") = "Вся площадь" Then
TmpLG("Prim1") = 0
TmpLG("PloLG") = TmpLG("Plo")
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = TmpLG("Prop")

TmpLG.UpdateBatch

'Если процент равен 0
If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
End If



                         End If
'-------------------------------------------------------

'------------------------- См на всех ------------------------------
              If TmpLG("Use") = "См на всех" Then
TmpLG("Prim1") = 0
If TmpLG("Plo") > TmpLG("Cocmin") Then TmpLG("PloLG") = TmpLG("Cocmin") Else TmpLG("PloLG") = TmpLG("Plo")
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = TmpLG("Prop")

TmpLG.UpdateBatch

'Если процент равен 0
If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
End If

                         End If
'-------------------------------------------------------

'------------------------- СМ на 1-го ------------------------------
              If TmpLG("Use") = "СМ на 1-го" Then
              


              
TmpLG("Prim1") = 0
If TmpLG("Plo") > DimSocmin(TmpLG("KodKat")) Then TmpLG("PloLG") = DimSocmin(TmpLG("KodKat")) Else TmpLG("PloLG") = TmpLG("Plo")

'MsgBox Str(DimSocmin(TmpLG("KodKat"))) + " " + TmpLG.Index

DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = 1

TmpLG.UpdateBatch


'Если процент равен 0
If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
End If

                         End If
'-------------------------------------------------------
'------------------------- См/кол.жил ------------------------------
              If TmpLG("Use") = "См/кол.жил" Then
TmpLG("Prim1") = 0
If TmpLG("Prop") <> 0 Then
If TmpLG("Plo") > TmpLG("Cocmin") / TmpLG("Prop") Then TmpLG("PloLG") = TmpLG("Cocmin") / TmpLG("Prop") Else TmpLG("PloLG") = TmpLG("Plo")
Else
TmpLG("PloLG") = 0
End If
If Not TmpLG("OtheCode") Then DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = 1

TmpLG.UpdateBatch
                         End If
'-------------------------------------------------------

'------------------------- На одного ------------------------------
              If TmpLG("Use") = "На одного" Then
TmpLG("Prim1") = 0
If TmpLG("Plo") > TmpLG("Plo") / TmpLG("Prop") Then TmpLG("PloLG") = TmpLG("Plo") / TmpLG("Prop") Else TmpLG("PloLG") = TmpLG("Plo")

' Добавлено для уменьшения лиготируемой площади до соцминимума
'If TmpLG("PloLG") > DimSocmin(TmpLG("KodKat")) Then TmpLG("PloLG") = DimSocmin(TmpLG("KodKat"))

DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = 1

TmpLG.UpdateBatch
                         End If
'-------------------------------------------------------

TmpLG.MoveNext
'en:
Loop


'**********************************************************************************
'***********************Теперь убераем льготы ДВОЙНИКИ*****************************
'**********************************************************************************
Itog = 0
PromItog = 0
LastIn = 0



For I = 1 To TmpLG.RecordCount
Dbl = DBLCode(I)
For j = 1 To TmpLG.RecordCount

If Dbl = DBLCode(j) And I <> j Then

TmpLG.AbsolutePosition = j

' Для площади
If TmpLG("Use") = "Вся площадь" Or TmpLG("Use") = "См на всех" Or TmpLG("Use") = "СМ на 1-го" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "См/кол.жил" Then PromItog = TmpLG("PloLG") * TmpLG("Procent") / 100

' Для количества прописанных
If TmpLG("Use") = "На 1-го" Or TmpLG("Use") = "На всех" Then PromItog = TmpLG("Prim1") * TmpLG("Procent") / 100

If Itog <= PromItog Then
Itog = PromItog
' Возвращаемся назад и проставляем -1 в Prim ЛЬГОТА НЕ ПРИМЕНЯЕТСЯ
'If LastIn <> 0 Then TmpLG.AbsolutePosition = LastIn

TmpLG.AbsolutePosition = I
TmpLG("Prim") = 1
TmpLG.UpdateBatch

If LastIn <> 0 Then
TmpLG.AbsolutePosition = LastIn
TmpLG("Prim") = -1
TmpLG.UpdateBatch
End If

LastIn = j
Else
TmpLG("Prim") = -1
TmpLG.UpdateBatch
End If

'MsgBox "Повтор" + "  " + DBLCode(j) + " " + Str(TmpLG("KodKat")) + " " + Str(Itog) + " " + TmpLG("Use") + " " + "Позиция=" + Str(TmpLG.AbsolutePosition)

'
'if tmplg("")
'
End If

Next j
Next I

TmpLG.Close

'**********************************************************************************
'***********************ВЫБЕРАЕМ НАИЛУЧШУЮ ЛЬГОТУ ДЛЯ ПРИМЕНЯЕМЫХ******************
'**********************************************************************************

' Список одновременно применяемых льгот ПЛОЩАДЬ
' If TmpLG("Use")="На одного" or TmpLG("Use")="См/кол.жил" - то применяются совместно со всеми льготами
' Сначала расчитываем If TmpLG("Use")="Вся площадь" or TmpLG("Use")="См на всех" or TmpLG("Use")="СМ на 1-го"
' Потом "плюсуем" TmpLG("Use")="На одного" or TmpLG("Use")="См/кол.жил"



' Открываем TmpLG для применяемых льгот т.е. TmpLG("Prim")<>-1
' НЕ ИСКЛЮЧЕННЫЕ ЗАПИСИ Т.Е. Уже просчитаны, или еще не обработаны

TmpLG.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.Plo, tmp_lgota.Prop, tmp_lgota.Cocmin, tmp_lgota.OtheCode, tmp_lgota.parametr, tmp_lgota.itog, tmp_lgota.tarif, tmp_lgota.Itog1, tmp_lgota.Prim, tmp_lgota.PloLG, tmp_lgota.Key, tmp_lgota.Prim1, tmp_lgota.Koll, tmp_lgota.SovmPloLG, Adding.KodKat, Adding.dop FROM Adding RIGHT JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd WHERE (((tmp_lgota.UniKOd)=" + Выбор + ") AND ((tmp_lgota.Prim)>=0))"), Mconn, adOpenKeyset, adLockPessimistic

LastIn = 0

TmpLG.MoveFirst
Do While Not TmpLG.EOF
' Если есть льгота На всех то TmpLG("Prim") = 1 потом отбрасываем все остальные
If TmpLG("Use") = "На всех" Then
'If TmpLG("Use") = "На всех" Or TmpLG("Use") = "См на всех" Then
TmpLG("Prim") = 1
TmpLG("itog1") = 1 - TmpLG("Procent") / 100
TmpLG.UpdateBatch
LastIn = TmpLG.AbsolutePosition

'отбрасываем все остальные
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG.AbsolutePosition <> LastIn Then
TmpLG("Prim") = -10
TmpLG.UpdateBatch
End If
TmpLG.MoveNext
Loop
Exit Do
'Exit Sub
End If

TmpLG.MoveNext
Loop

' Если несколько TmpLG("Use") = "На 1-го"
TmpLG.Requery

KolL = 0
PSumm = 0
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "На 1-го" Then
'Itog - это процент льготы только по одной текущей льготе
'Itog1 - это ОБЩИЙ процент льготы по ВСЕМ льготам категории

'Подсчитываем количество таких льгот KolL
'и суммарный процент по льготе PSumm для последующего расчета ОБЩЕГО % по Л/сч
 
 KolL = KolL + 1
 P = (100 - TmpLG("Procent")) / 100
 

 If TmpLG("Prop") <> 0 Then
 TmpLG.Fields("itog").Value = (P + (TmpLG("Prop") - 1)) / TmpLG("Prop")
 PSumm = PSumm + 1 * P
 TmpLG("Prim") = 1
 TmpLG("PloLg") = 0
 TmpLG.UpdateBatch
 End If
 
 End If
TmpLG.MoveNext
Loop



'Теперь расчитываем общий процент льгот На 1-го для лиц счета и записываем в Itog1
TmpLG.Requery
ProcentLMAX = 0
TmpLG.MoveFirst
Do While Not TmpLG.EOF

If TmpLG("Use") = "На 1-го" Then
P = (100 - TmpLG("Procent")) / 100 '
If TmpLG("Prop") <> 0 Then ProcentLMAX = (PSumm + (TmpLG("Prop") - KolL)) / TmpLG("Prop")

TmpLG("itog1") = ProcentLMAX
TmpLG.UpdateBatch
End If

TmpLG.MoveNext
Loop


'**************************** Расчет от пложади***********************
TmpLG.Requery

'
TmpLG.MoveFirst

'***************"Вся площадь"**************

Do While Not TmpLG.EOF

If TmpLG("plo") <> 0 Then

If TmpLG("Use") = "Вся площадь" Then

TmpLG("SovmPloLG") = TmpLG("plo") * (100 - TmpLG("Procent")) / 100
TmpLG("Itog") = (100 - TmpLG("Procent")) / 100

If TmpLG("Procent") = 100 Then TmpLG("SovmPloLG") = TmpLG("plo")


TmpLG.UpdateBatch
End If
End If

TmpLG.MoveNext
Loop

'***************"См на всех"**************
TmpLG.MoveFirst
Do While Not TmpLG.EOF

If TmpLG("plo") <> 0 Then

If TmpLG("Use") = "См на всех" Then
TmpLG("SovmPloLG") = TmpLG("plolg") * (100 - TmpLG("Procent")) / 100
TmpLG("Itog") = (TmpLG("plolg") * (100 - TmpLG("Procent")) / 100) / TmpLG("plo")
TmpLG.UpdateBatch
End If
End If

TmpLG.MoveNext
Loop


'***************"См/кол.жил" "На одного" "СМ на 1-го" "СМ на 1-го"**************
SumLGPl = 0

TmpLG.MoveFirst
Ostat = TmpLG("plo")

Do While Not TmpLG.EOF

If TmpLG("plo") <> 0 Then

If TmpLG("Use") = "См/кол.жил" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "СМ на 1-го" Then
' Добавить or TmpLG("Use") = "СМ на 1-го"
'если надо учитывать совместно "См/кол.жил" "На одного" и "СМ на 1-го"

SumLGPl = SumLGPl + TmpLG("plolg")

If (TmpLG("plo") - SumLGPl) > 0 Then

Ostat = Ostat - TmpLG("plolg")

TmpLG("SovmPloLG") = TmpLG("plolg") * (100 - TmpLG("Procent")) / 100
TmpLG("Itog") = (TmpLG("plolg") * (100 - TmpLG("Procent")) / 100) / TmpLG("plo")
PloSumm = PloSumm + TmpLG("plolg") * (100 - TmpLG("Procent")) / 100
TmpLG.UpdateBatch
Else
'TmpLG("SovmPloLG") = (TmpLG("plo") - SumLGPl) * -1 * (100 - TmpLG("Procent")) / 100
TmpLG("plolg") = Ostat
TmpLG("Itog") = (Ostat * (100 - TmpLG("Procent")) / 100) / TmpLG("plo")

PloSumm = PloSumm + Ostat * (100 - TmpLG("Procent")) / 100

TmpLG.UpdateBatch
Exit Do
End If
End If
End If

TmpLG.MoveNext
Loop


'Теперь проставляем суммарную площадь для всех совместно применяемых льгот
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("plo") <> 0 Then
If TmpLG("Use") = "См/кол.жил" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "СМ на 1-го" Then
' Добавить or TmpLG("Use") = "СМ на 1-го"
'если надо учитывать совместно "См/кол.жил" "На одного" и "СМ на 1-го"

TmpLG("SovmPloLG") = PloSumm
TmpLG.UpdateBatch
End If
End If
TmpLG.MoveNext
Loop


' Теперь выбираем лучшую льготу, это те льготы, которые имеют MAX SovmPloLG
' Этих льгот может быть несколько
Max = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If Max < TmpLG("SovmPloLG") Then
Max = TmpLG("SovmPloLG")
End If
TmpLG.MoveNext
Loop

'Проставляем Prim=1 Prim=-1
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("SovmPloLG") = Max Then TmpLG("Prim") = 1 Else TmpLG("Prim") = -1
TmpLG.UpdateBatch
TmpLG.MoveNext
Loop



TmpLG.Requery
TmpLG.MoveFirst
I = 0
j = 0
Do While Not TmpLG.EOF
If TmpLG("Use") = "См на всех" Or TmpLG("Use") = "Вся площадь" Then I = 1
If TmpLG("Use") = "См/кол.жил" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "СМ на 1-го" Then j = 1
TmpLG.MoveNext
Loop

If I = 1 And j = 1 Then

TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "См на всех" Or TmpLG("Use") = "Вся площадь" Then TmpLG("Prim") = 1
If TmpLG("Use") = "См/кол.жил" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "СМ на 1-го" Then TmpLG("Prim") = -1
TmpLG.UpdateBatch
TmpLG.MoveNext
Loop
End If

' Теперь расчитываем ОБЩИЙ процент и проставляем в Adding
TmpLG.Requery
TmpLG.MoveFirst
strLg = " "
Do While Not TmpLG.EOF
strLg = strLg + Str(TmpLG("KodKls")) + ","
If TmpLG("Use") = "См на всех" Or TmpLG("Use") = "Вся площадь" Or TmpLG("Use") = "См/кол.жил" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "СМ на 1-го" Then



TmpLG("Itog1") = (TmpLG("plo") - TmpLG("SovmPloLG")) / TmpLG("plo")
TmpLG.UpdateBatch
End If
TmpLG.MoveNext
Loop
strLg = Trim(strLg)

Mconn.Execute ("UPDATE Adding INNER JOIN TMP_Lgota ON Adding.Key = TMP_Lgota.UniKOd SET Adding.LgotaP = [TMP_Lgota]![Itog1], Adding.LgotaKod = " + Chr(34) + strLg + Chr(34) + " WHERE (((Adding.Key)=" + Выбор + ") AND ((TMP_Lgota.Prim)=1))")

''' ИСПРАВЛЕНИЕ ОШИБОК И НЕДАЧЕТОВ ВЫЯВЛЕННЫХ ПРИ ЭКСПЛУАТАЦИИ СИСТЕМЫ

' 1 Если встречаются 2 и более льготы "Вся площадь" то отбрасываем все кроме первой
n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "Вся площадь" Then
n = n + 1
If n > 1 Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop

'2. Если встречаются 2 и более льготы "СМ на всех" привышающие общую площадь то отбрасываем все кроме первой

n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "См на всех" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop

'3. Если встречаются 2 и более льготы "На одного" привышающие общую площадь то отбрасываем все кроме первой

n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "На одного" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop


'4. Если встречаются 2 и более льготы "СМ/кол.жил" привышающие общую площадь то отбрасываем все кроме первой

n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "См/кол.жил" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop

' 5. Проверяем не привысила ли лиготируемая площадь общую


n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "См/кол.жил" Or TmpLG("Use") = "На одного" Or TmpLG("Use") = "СМ на 1-го" Or TmpLG("Use") = "Вся площадь" Or TmpLG("Use") = "См на всех" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("plolg") = TmpLG("plolg") - (n - TmpLG("plo"))
If TmpLG("plolg") = 0 Then
TmpLG("prim") = -1
End If
TmpLG.UpdateBatch
End If
End If
TmpLG.MoveNext
Loop

End Sub
