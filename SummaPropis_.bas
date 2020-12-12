Attribute VB_Name = "SummaPropis_"
' Представление числа прописью на русском языке.
' Поддержка чисел типа Currency во всем диапазоне (т.е до ~922 триллионов рублей)
' При втором аргументе функции равном 0, вывод только числа прописью,
' при втором аргументе функции равном 1, дополнительно вывод рублей и копеек


Private Skl As Byte

Public Function NumStr(n As Currency, Optional rub As Boolean) As String
Dim s As String, R As String, K As String
Dim t, u, v, w As Integer

s = ""

If n < 0 Then
n = Abs(n)
s = "минус"
End If
'-----------------------------------------------------------------------------
v = (n - Fix(n)) * 100 ' Число копеек
w = Val(Right(Format(v), 1)) ' Получить число единиц копеек

n = Fix(n) ' Целое число рублей
t = Val(Right(Format(n), 2)) ' Получить две последние цифры рублей
u = Val(Right(t, 1)) ' Получить число единиц рублей

If t > 10 And t < 15 Then
R = " рублей" ' Получить подпись для рублей
ElseIf u = 1 Then
R = " рубль"
ElseIf u > 1 And u < 5 Then
R = " рубля"
Else
R = " рублей"
End If

If v > 10 And v < 15 Then
K = " копеек." ' Получить подпись для копеек
ElseIf w = 1 Then
K = " копейка."
ElseIf w > 1 And w < 5 Then
K = " копейки."
Else
K = " копеек."
End If

'-----------------------------------------------------------------------------
If n > 1000000000000# Then
s = AddStr(s, NumStr2(Int(n / 1000000000000#), True))
Select Case Skl
Case 0
s = AddStr(s, "триллион")
Case 1
s = AddStr(s, "триллиона")
Case 2
s = AddStr(s, "триллионов")
End Select
n = n - Int(n / 1000000000000#) * 1000000000000#
End If

If n > 1000000000 Then
s = AddStr(s, NumStr2(Int(n / 1000000000), True))
Select Case Skl
Case 0
s = AddStr(s, "миллиард")
Case 1
s = AddStr(s, "миллиарда")
Case 2
s = AddStr(s, "миллиардов")
End Select
n = n - Int(n / 1000000000) * 1000000000
End If

If n > 1000000 Then
s = AddStr(s, NumStr2(n \ 1000000, True))
Select Case Skl
Case 0
s = AddStr(s, "миллион")
Case 1
s = AddStr(s, "миллиона")
Case 2
s = AddStr(s, "миллионов")
End Select
n = n Mod 1000000
End If

If n > 1000 Then
s = AddStr(s, NumStr2(n \ 1000, False))
Select Case Skl
Case 0
s = AddStr(s, "тысяча")
Case 1
s = AddStr(s, "тысячи")
Case 2
s = AddStr(s, "тысяч")
End Select
n = n Mod 1000
End If

If n > 0 Then
s = AddStr(s, NumStr2(n, True))
End If

If s = "" Then
s = "ноль"
ElseIf s = "минус" Then
s = s + " ноль"
End If

NumStr = StrConv(Mid(s, 1, 1), vbUpperCase) + Mid(s, 2, Len(s) - 1)
If (rub) Then NumStr = NumStr & R & Format(v, " 00") & K

End Function
'-----------------------------------------------------------------------------

Private Function NumStr2(n As Currency, male As Boolean) As String
Dim s As String
s = ""
If n >= 100 Then
s = NumStr1(((n \ 100) * 100), male)
n = n Mod 100
End If
If n >= 20 Then
s = AddStr(s, NumStr1(((n \ 10) * 10), male))
n = n Mod 10
End If
NumStr2 = AddStr(s, NumStr1(n, male))
End Function
'-----------------------------------------------------------------------------

Private Function NumStr1(n As Currency, male As Boolean) As String
Skl = 2
Select Case n
Case 100
NumStr1 = "сто"
Case 200
NumStr1 = "двести"
Case 300
NumStr1 = "триста"
Case 400
NumStr1 = "четыреста"
Case 500
NumStr1 = "пятьсот"
Case 600
NumStr1 = "шестьсот"
Case 700
NumStr1 = "семьсот"
Case 800
NumStr1 = "восемьсот"
Case 900
NumStr1 = "девятьсот"
Case 11
NumStr1 = "одиннадцать"
Case 12
NumStr1 = "двенадцать"
Case 13
NumStr1 = "тринадцать"
Case 14
NumStr1 = "четырнадцать"
Case 15
NumStr1 = "пятнадцать"
Case 16
NumStr1 = "шестнадцать"
Case 17
NumStr1 = "семнадцать"
Case 18
NumStr1 = "восемнадцать"
Case 19
NumStr1 = "девятнадцать"
Case 20
NumStr1 = "двадцать"
Case 30
NumStr1 = "тридцать"
Case 40
NumStr1 = "сорок"
Case 50
NumStr1 = "пятьдесят"
Case 60
NumStr1 = "шестьдесят"
Case 70
NumStr1 = "семьдесят"
Case 80
NumStr1 = "восемьдесят"
Case 90
NumStr1 = "девяносто"
Case 1
Skl = 0
If male Then
NumStr1 = "один"
Else
NumStr1 = "одна"
End If
Case 2
Skl = 1
If male Then
NumStr1 = "два"
Else
NumStr1 = "две"
End If
Case 3
Skl = 1
NumStr1 = "три"
Case 4
Skl = 1
NumStr1 = "четыре"
Case 5
NumStr1 = "пять"
Case 6
NumStr1 = "шесть"
Case 7
NumStr1 = "семь"
Case 8
NumStr1 = "восемь"
Case 9
NumStr1 = "девять"
Case 10
NumStr1 = "десять"
End Select
End Function
'-----------------------------------------------------------------------------

Private Function AddStr(S1 As String, S2 As String)
If S1 = "" Then
AddStr = S2
ElseIf S2 = "" Then
AddStr = S1
Else
AddStr = S1 + " " + S2
End If
End Function

