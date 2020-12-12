Attribute VB_Name = "Propis"
'Attribute VB_Name = "Propis"
'Option Compare Database

Private Function sumPropis(dSumma As Double) As String
Dim sp As String ' ������ ��������
Dim sn As String ' �������� ������������� �����
Dim sd As String ' ���������� �������
Dim rub(10) As String ' ����� ������
Dim mlrd(10) As String ' ����� ����������
Dim mln(10) As String ' ����� ���������
Dim tys(10) As String ' ����� �����

rub(1) = " ����� "
rub(2) = " ����� "
rub(3) = " ����� "
rub(4) = " ����� "
rub(5) = " ������ "
rub(6) = " ������ "
rub(7) = " ������ "
rub(8) = " ������ "
rub(9) = " ������ "
rub(0) = " ������ "
'
tys(1) = " ������ "
tys(2) = " ������ "
tys(3) = " ������ "
tys(4) = " ������ "
tys(5) = " ����� "
tys(6) = " ����� "
tys(7) = " ����� "
tys(8) = " ����� "
tys(9) = " ����� "
tys(0) = " ����� "
'
mln(1) = " ������� "
mln(2) = " �������� "
mln(3) = " �������� "
mln(4) = " �������� "
mln(5) = " ��������� "
mln(6) = " ��������� "
mln(7) = " ��������� "
mln(8) = " ��������� "
mln(9) = " ��������� "
mln(0) = " ��������� "
'
mlrd(1) = " �������� "
mlrd(2) = " ��������� "
mlrd(3) = " ��������� "
mlrd(4) = " ��������� "
mlrd(5) = " ���������� "
mlrd(6) = " ���������� "
mlrd(7) = " ���������� "
mlrd(8) = " ���������� "
mlrd(9) = " ���������� "
mlrd(0) = " ���������� "
'
'�������������
Let sumPropis = ""
'��������� ����� �� ������������
If dSumma <= 0 Then Exit Function
'��������� �� �������
sn = Format(Int(dSumma), "000000000000")
sd = Format(Round((dSumma - Val(sn)) * 100, 0), "00")
'���������������� ������
'��������� - ����� ����� ����������
If Val(Mid(sn, 1, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 1, 3), False) & IIf(Mid(sn, 2, 1) = 1, mlrd(0), mlrd(Mid(sn, 3, 1)))
'��������
If Val(Mid(sn, 4, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 4, 3), False) & IIf(Mid(sn, 5, 1) = 1, mln(0), mln(Mid(sn, 6, 1)))
'������
If Val(Mid(sn, 7, 3)) <> 0 Then sumPropis = sumPropis & sTriple(Mid(sn, 7, 3), True) & IIf(Mid(sn, 8, 1) = 1, tys(0), tys(Mid(sn, 9, 1)))
'� �������
sumPropis = sumPropis & sTriple(Mid(sn, 10, 3), False)
'���������� ���������
sumPropis = sumPropis & IIf(Mid(sn, 11, 1) = 1, rub(0), rub(Right(sn, 1))) & sd & " ���."
'
End Function

Private Function sTriple(sRazr As String, bGender As Boolean) As String
'������� ��������� ����������� ����� � ����� �������� � ������ ����
Dim Ed(20) As String  ' ������ ������
Dim des(10) As String ' ������ �������
Dim sot(10) As String ' ������ �����
'�������� ������
Ed(0) = ""
Ed(1) = " ����"
Ed(2) = " ���"
Ed(3) = " ���"
Ed(4) = " ������"
Ed(5) = " ����"
Ed(6) = " �����"
Ed(7) = " ����"
Ed(8) = " ������"
Ed(9) = " ������"
Ed(10) = " ������"
Ed(11) = " �����������"
Ed(12) = " ����������"
Ed(13) = " ����������"
Ed(14) = " ������������"
Ed(15) = " ����������"
Ed(16) = " �����������"
Ed(17) = " ����������"
Ed(18) = " ������������"
Ed(19) = " ������������"
'�������� ��������
des(0) = ""
des(1) = " ������"
des(2) = " ��������"
des(3) = " ��������"
des(4) = " �����"
des(5) = " ���������"
des(6) = " ����������"
des(7) = " ���������"
des(8) = " �����������"
des(9) = " ���������"
'�������� �����
sot(0) = ""
sot(1) = " ���"
sot(2) = " ������"
sot(3) = " ������"
sot(4) = " ���������"
sot(5) = " �������"
sot(6) = " ��������"
sot(7) = " �������"
sot(8) = " ���������"
sot(9) = " ���������"
' ���� ���� ��� �����
If bGender Then
    Ed(1) = " ����"
    Ed(2) = " ���"
End If
' ���������� � �������
sTriple = sTriple & sot(Mid(sRazr, 1, 1))
' ���� ������� �������
If Mid(sRazr, 2, 2) > 10 And Mid(sRazr, 2, 2) < 20 Then
    sTriple = sTriple & Ed(Mid(sRazr, 2, 2))
Else
' ����� ������ - ���� ������� �� ������
    sTriple = sTriple & des(Mid(sRazr, 2, 1))
    sTriple = sTriple & Ed(Mid(sRazr, 3, 1))
End If

End Function



