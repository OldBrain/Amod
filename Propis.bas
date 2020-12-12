Attribute VB_Name = "Propis"
'Attribute VB_Name = "Propis"
'Option Compare Database

Private Function sumPropis(dSumma As Double) As String
Dim sp As String ' строка прописью
Dim sn As String ' строчное представление числа
Dim sd As String ' количество дробное
Dim rub(10) As String ' имена валюты
Dim mlrd(10) As String ' имена миллиардов
Dim mln(10) As String ' имена миллионов
Dim tys(10) As String ' имена тысяч

rub(1) = " рубль "
rub(2) = " рубля "
rub(3) = " рубля "
rub(4) = " рубля "
rub(5) = " рублей "
rub(6) = " рублей "
rub(7) = " рублей "
rub(8) = " рублей "
rub(9) = " рублей "
rub(0) = " рублей "
'
tys(1) = " тысяча "
tys(2) = " тысячи "
tys(3) = " тысячи "
tys(4) = " тысячи "
tys(5) = " тысяч "
tys(6) = " тысяч "
tys(7) = " тысяч "
tys(8) = " тысяч "
tys(9) = " тысяч "
tys(0) = " тысяч "
'
mln(1) = " миллион "
mln(2) = " миллиона "
mln(3) = " миллиона "
mln(4) = " миллиона "
mln(5) = " миллионов "
mln(6) = " миллионов "
mln(7) = " миллионов "
mln(8) = " миллионов "
mln(9) = " миллионов "
mln(0) = " миллионов "
'
mlrd(1) = " миллиард "
mlrd(2) = " миллиарда "
mlrd(3) = " миллиарда "
mlrd(4) = " миллиарда "
mlrd(5) = " миллиардов "
mlrd(6) = " миллиардов "
mlrd(7) = " миллиардов "
mlrd(8) = " миллиардов "
mlrd(9) = " миллиардов "
mlrd(0) = " миллиардов "
'
'инициализация
Let sumPropis = ""
'проверить число на правильность
If dSumma <= 0 Then Exit Function
'разложить по тройкам
sn = Format(Int(dSumma), "000000000000")
sd = Format(Round((dSumma - Val(sn)) * 100, 0), "00")
'проанализировать тройки
'миллиарды - авось когда пригодятся
If Val(Mid(sn, 1, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 1, 3), False) & IIf(Mid(sn, 2, 1) = 1, mlrd(0), mlrd(Mid(sn, 3, 1)))
'миллионы
If Val(Mid(sn, 4, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 4, 3), False) & IIf(Mid(sn, 5, 1) = 1, mln(0), mln(Mid(sn, 6, 1)))
'тысячи
If Val(Mid(sn, 7, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 7, 3), True) & IIf(Mid(sn, 8, 1) = 1, tys(0), tys(Mid(sn, 9, 1)))
'и единицы
sumPropis = sumPropis & sTriple(Mid(sn, 10, 3), False)
'возвратить результат
sumPropis = sumPropis & IIf(Mid(sn, 11, 1) = 1, rub(0), rub(Right(sn, 1))) & sd & " коп."
'
End Function

Private Function sTriple(sRazr As String, bGender As Boolean) As String
'Функция переводит трехзначное число в число прописью с учетом рода
Dim Ed(20) As String  ' массив единиц
Dim des(10) As String ' массив десяток
Dim sot(10) As String ' массив сотен
'значения единиц
Ed(0) = ""
Ed(1) = " один"
Ed(2) = " два"
Ed(3) = " три"
Ed(4) = " четыре"
Ed(5) = " пять"
Ed(6) = " шесть"
Ed(7) = " семь"
Ed(8) = " восемь"
Ed(9) = " девять"
Ed(10) = " десять"
Ed(11) = " одиннадцать"
Ed(12) = " двенадцать"
Ed(13) = " тринадцать"
Ed(14) = " четырнадцать"
Ed(15) = " пятнадцать"
Ed(16) = " шестнадцать"
Ed(17) = " семнадцать"
Ed(18) = " восемнадцать"
Ed(19) = " девятнадцать"
'значения десятков
des(0) = ""
des(1) = " десять"
des(2) = " двадцать"
des(3) = " тридцать"
des(4) = " сорок"
des(5) = " пятьдесят"
des(6) = " шестьдесят"
des(7) = " семьдесят"
des(8) = " восемьдесят"
des(9) = " девяносто"
'значения сотен
sot(0) = ""
sot(1) = " сто"
sot(2) = " двести"
sot(3) = " триста"
sot(4) = " четыреста"
sot(5) = " пятьсот"
sot(6) = " шестьсот"
sot(7) = " семьсот"
sot(8) = " восемьсот"
sot(9) = " девятьсот"
' учет рода для тысяч
If bGender Then
    Ed(1) = " одна"
    Ed(2) = " две"
End If
' трансляция в пропись
sTriple = sTriple & sot(Mid(sRazr, 1, 1))
' учет первого десятка
If Mid(sRazr, 2, 2) > 10 And Mid(sRazr, 2, 2) < 20 Then
    sTriple = sTriple & Ed(Mid(sRazr, 2, 2))
Else
' общий случай - если десятка не первая
    sTriple = sTriple & des(Mid(sRazr, 2, 1))
    sTriple = sTriple & Ed(Mid(sRazr, 3, 1))
End If

End Function



