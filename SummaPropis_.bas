Attribute VB_Name = "SummaPropis_"
' ������������� ����� �������� �� ������� �����.
' ��������� ����� ���� Currency �� ���� ��������� (�.� �� ~922 ���������� ������)
' ��� ������ ��������� ������� ������ 0, ����� ������ ����� ��������,
' ��� ������ ��������� ������� ������ 1, ������������� ����� ������ � ������


Private Skl As Byte

Public Function NumStr(n As Currency, Optional rub As Boolean) As String
Dim s As String, R As String, K As String
Dim t, u, v, w As Integer

s = ""

If n < 0 Then
n = Abs(n)
s = "�����"
End If
'-----------------------------------------------------------------------------
v = (n - Fix(n)) * 100 ' ����� ������
w = Val(Right(Format(v), 1)) ' �������� ����� ������ ������

n = Fix(n) ' ����� ����� ������
t = Val(Right(Format(n), 2)) ' �������� ��� ��������� ����� ������
u = Val(Right(t, 1)) ' �������� ����� ������ ������

If t > 10 And t < 15 Then
R = " ������" ' �������� ������� ��� ������
ElseIf u = 1 Then
R = " �����"
ElseIf u > 1 And u < 5 Then
R = " �����"
Else
R = " ������"
End If

If v > 10 And v < 15 Then
K = " ������." ' �������� ������� ��� ������
ElseIf w = 1 Then
K = " �������."
ElseIf w > 1 And w < 5 Then
K = " �������."
Else
K = " ������."
End If

'-----------------------------------------------------------------------------
If n > 1000000000000# Then
s = AddStr(s, NumStr2(Int(n / 1000000000000#), True))
Select Case Skl
Case 0
s = AddStr(s, "��������")
Case 1
s = AddStr(s, "���������")
Case 2
s = AddStr(s, "����������")
End Select
n = n - Int(n / 1000000000000#) * 1000000000000#
End If

If n > 1000000000 Then
s = AddStr(s, NumStr2(Int(n / 1000000000), True))
Select Case Skl
Case 0
s = AddStr(s, "��������")
Case 1
s = AddStr(s, "���������")
Case 2
s = AddStr(s, "����������")
End Select
n = n - Int(n / 1000000000) * 1000000000
End If

If n > 1000000 Then
s = AddStr(s, NumStr2(n \ 1000000, True))
Select Case Skl
Case 0
s = AddStr(s, "�������")
Case 1
s = AddStr(s, "��������")
Case 2
s = AddStr(s, "���������")
End Select
n = n Mod 1000000
End If

If n > 1000 Then
s = AddStr(s, NumStr2(n \ 1000, False))
Select Case Skl
Case 0
s = AddStr(s, "������")
Case 1
s = AddStr(s, "������")
Case 2
s = AddStr(s, "�����")
End Select
n = n Mod 1000
End If

If n > 0 Then
s = AddStr(s, NumStr2(n, True))
End If

If s = "" Then
s = "����"
ElseIf s = "�����" Then
s = s + " ����"
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
NumStr1 = "���"
Case 200
NumStr1 = "������"
Case 300
NumStr1 = "������"
Case 400
NumStr1 = "���������"
Case 500
NumStr1 = "�������"
Case 600
NumStr1 = "��������"
Case 700
NumStr1 = "�������"
Case 800
NumStr1 = "���������"
Case 900
NumStr1 = "���������"
Case 11
NumStr1 = "�����������"
Case 12
NumStr1 = "����������"
Case 13
NumStr1 = "����������"
Case 14
NumStr1 = "������������"
Case 15
NumStr1 = "����������"
Case 16
NumStr1 = "�����������"
Case 17
NumStr1 = "����������"
Case 18
NumStr1 = "������������"
Case 19
NumStr1 = "������������"
Case 20
NumStr1 = "��������"
Case 30
NumStr1 = "��������"
Case 40
NumStr1 = "�����"
Case 50
NumStr1 = "���������"
Case 60
NumStr1 = "����������"
Case 70
NumStr1 = "���������"
Case 80
NumStr1 = "�����������"
Case 90
NumStr1 = "���������"
Case 1
Skl = 0
If male Then
NumStr1 = "����"
Else
NumStr1 = "����"
End If
Case 2
Skl = 1
If male Then
NumStr1 = "���"
Else
NumStr1 = "���"
End If
Case 3
Skl = 1
NumStr1 = "���"
Case 4
Skl = 1
NumStr1 = "������"
Case 5
NumStr1 = "����"
Case 6
NumStr1 = "�����"
Case 7
NumStr1 = "����"
Case 8
NumStr1 = "������"
Case 9
NumStr1 = "������"
Case 10
NumStr1 = "������"
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

