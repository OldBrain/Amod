Attribute VB_Name = "Rashet"
Option Explicit
Dim ProcentLMAX As Double ' ��� �������� ������ �������� ����� ��� ��� ����� /������������� ����/
Dim KolL As Integer ' ��� ������� ���������� ���������� /������������� ����/
Dim P As Double '/������������� ����/ ��� %
Dim PSumm As Double '/������������� ����/ ��� ���������� %
Dim PloSumm As Double '/������������� ����/ ��� �������� ��������� ������������ �������
Dim Ostat As Double
Dim SumLGPl As Double
Dim strLg As String
Dim Max As Double '/������������� ����/ ��� ����������� ������������(������) ������
Dim I As Integer
Dim j As Integer
Dim LastIn As Integer ' ��� ����������� ����������� ������ ������ ���������� ������� �����
Dim Dbl As Double
Dim PromItog As Double
Dim DBLCode(100) As String ' ������ ��� ���������  TmpLG("OtheCode")
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
Dim rsSocmin As ADODB.Recordset ' ��� ����������� ����������� �� 1-��
Dim TmpLG As ADODB.Recordset ' ��� ������ �� ������ ��������� �������
'Dim rsTmpDBL As ADODB.Recordset ' ��� ������� �����

Public Sub ������(����� As String)

'DoEvents







PloSumm = 0

'�������� ��������� �������
Mconn.Execute ("UPDATE TMP_Lgota SET TMP_Lgota.Prim = 0, TMP_Lgota.PloLG = 0, TMP_Lgota.Prim1 = 0, TMP_Lgota.itog1 = 1, TMP_Lgota.SovmPloLG = 0 WHERE (((TMP_Lgota.UniKOd)=" + ����� + "))")

Set TmpLG = New ADODB.Recordset
'Set rsTmpDBL = New ADODB.Recordset

TmpLG.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.Plo, tmp_lgota.Prop, tmp_lgota.Cocmin, tmp_lgota.OtheCode, tmp_lgota.parametr, tmp_lgota.itog, tmp_lgota.tarif, tmp_lgota.Itog1, tmp_lgota.Prim, tmp_lgota.PloLG, tmp_lgota.Key, tmp_lgota.Prim1, tmp_lgota.Koll, tmp_lgota.SovmPloLG, Adding.KodKat, Adding.dop FROM Adding RIGHT JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd WHERE (((tmp_lgota.UniKOd)=" + ����� + "))"), Mconn, adOpenKeyset, adLockPessimistic

'���������� ���������� �� 1-�� ��� ������ ��������� �������
Set rsSocmin = New ADODB.Recordset
rsSocmin.Open ("SELECT Socmin.KodKategor, Socmin.koli, Socmin.Value From Socmin WHERE (((Socmin.koli)=1))"), Mconn

rsSocmin.MoveFirst
Do While Not rsSocmin.EOF
DimSocmin(rsSocmin("KodKategor")) = rsSocmin("Value") + TmpLG("Dop")
rsSocmin.MoveNext
Loop






' ******* ���������� ��� ������ ����� ��� �������

' �� ��������� � TmpLG("Prim")=0 - ������ ��� �� ��������������
'                TmpLG("Prim")=-1 - ������ �� �����������
'                TmpLG("Prim")=1 - ������ �����������



'                                   ��� ���� �����������

'                   ��� ���������� ������������� �� ���������� ������� ���������� ������

' 1. ���� TmpLG("Use")="�� 1-��"  ��� TmpLG("Use")="�� ����" �� ��� ������ �� ���������� �����������
'    ������ �������� ���������� ��� ������� ������ ����� �������� TmpLG("Prim1").
' 2. ���� TmpLG("Use")="�� 1-��" �� Prim1=1, ���� TmpLG("Use")="�� ����" �� Prim1=Prop (���-�� �����������)
' 3. ����� ���� �������� TmpLG("OtheCode") ����� ���������� �������� �� ��������� ���������� �����
'    ������ � ������� ��������� Prim1*Procent
' 4. ���� ������ ����������� �� ����������� TmpLG("Prim")=1 ����� TmpLG("Prim")=0

'                    ������������ ��������� ������� ������ ��� ���, �����
' 1. ���� TmpLG("Prim")=0 �� ����������
' 2. ���� TmpLG("Prim")=1 �� ������� ?????


' ����������� �������� Prim1(���-�� ���.�� ������� ����������� ������)
' � �������� PloLg

'If TmpLG.RecordCount = 0 Then Exit Sub
'************* ���� ��� ������� � ������� � ����� TmpLgota to �����****
On Error GoTo en
en:
If Err.Number = 3021 Then
'MsgBox Err.Description
Err.Clear
strLg = "0"
Mconn.Execute ("UPDATE Adding SET Adding.LgotaP =1, Adding.LgotaKod = " + Chr(34) + strLg + Chr(34) + " WHERE (((Adding.Key)=" + ����� + "))")
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


'���� ��� ����������� ��� ����� ������� �� �����
If TmpLG("Plo") = 0 And TmpLG("Use") <> "�� 1-��" And TmpLG("Use") <> "�� ����" Then
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
'MsgBox Str(TmpLG.AbsolutePosition) + " " + �����


'���� ��� ������ �� �����
If TmpLG("Tarif") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch

Exit Sub
End If






'------------------------- �� 1-�� ------------------------------
              If TmpLG("Use") = "�� 1-��" Then
TmpLG("Prim1") = 1
TmpLG("PloLG") = 0
PloSumm = 0
If Not TmpLG("OtheCode") Then DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")
TmpLG("Koll") = 1
TmpLG.UpdateBatch
                         End If
'-------------------------------------------------------

'------------------------- �� ���� ------------------------------
              If TmpLG("Use") = "�� ����" Then
TmpLG("Prim1") = TmpLG("Prop")
TmpLG("PloLG") = 0
TmpLG("Koll") = TmpLG("Prop")

TmpLG.UpdateBatch
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

                         End If
'-------------------------------------------------------

'------------------------- ��� ������� ------------------------------
              If TmpLG("Use") = "��� �������" Then
TmpLG("Prim1") = 0
TmpLG("PloLG") = TmpLG("Plo")
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = TmpLG("Prop")

TmpLG.UpdateBatch

'���� ������� ����� 0
If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
End If



                         End If
'-------------------------------------------------------

'------------------------- �� �� ���� ------------------------------
              If TmpLG("Use") = "�� �� ����" Then
TmpLG("Prim1") = 0
If TmpLG("Plo") > TmpLG("Cocmin") Then TmpLG("PloLG") = TmpLG("Cocmin") Else TmpLG("PloLG") = TmpLG("Plo")
DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = TmpLG("Prop")

TmpLG.UpdateBatch

'���� ������� ����� 0
If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
End If

                         End If
'-------------------------------------------------------

'------------------------- �� �� 1-�� ------------------------------
              If TmpLG("Use") = "�� �� 1-��" Then
              


              
TmpLG("Prim1") = 0
If TmpLG("Plo") > DimSocmin(TmpLG("KodKat")) Then TmpLG("PloLG") = DimSocmin(TmpLG("KodKat")) Else TmpLG("PloLG") = TmpLG("Plo")

'MsgBox Str(DimSocmin(TmpLG("KodKat"))) + " " + TmpLG.Index

DBLCode(TmpLG.AbsolutePosition) = TmpLG("OtheCode")

TmpLG("Koll") = 1

TmpLG.UpdateBatch


'���� ������� ����� 0
If TmpLG("Procent") = 0 Then
TmpLG("PloLG") = 0
TmpLG("Prim1") = -1
TmpLG("Prim") = -1
TmpLG("Koll") = 0
TmpLG.UpdateBatch
End If

                         End If
'-------------------------------------------------------
'------------------------- ��/���.��� ------------------------------
              If TmpLG("Use") = "��/���.���" Then
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

'------------------------- �� ������ ------------------------------
              If TmpLG("Use") = "�� ������" Then
TmpLG("Prim1") = 0
If TmpLG("Plo") > TmpLG("Plo") / TmpLG("Prop") Then TmpLG("PloLG") = TmpLG("Plo") / TmpLG("Prop") Else TmpLG("PloLG") = TmpLG("Plo")

' ��������� ��� ���������� ������������ ������� �� �����������
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
'***********************������ ������� ������ ��������*****************************
'**********************************************************************************
Itog = 0
PromItog = 0
LastIn = 0



For I = 1 To TmpLG.RecordCount
Dbl = DBLCode(I)
For j = 1 To TmpLG.RecordCount

If Dbl = DBLCode(j) And I <> j Then

TmpLG.AbsolutePosition = j

' ��� �������
If TmpLG("Use") = "��� �������" Or TmpLG("Use") = "�� �� ����" Or TmpLG("Use") = "�� �� 1-��" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "��/���.���" Then PromItog = TmpLG("PloLG") * TmpLG("Procent") / 100

' ��� ���������� �����������
If TmpLG("Use") = "�� 1-��" Or TmpLG("Use") = "�� ����" Then PromItog = TmpLG("Prim1") * TmpLG("Procent") / 100

If Itog <= PromItog Then
Itog = PromItog
' ������������ ����� � ����������� -1 � Prim ������ �� �����������
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

'MsgBox "������" + "  " + DBLCode(j) + " " + Str(TmpLG("KodKat")) + " " + Str(Itog) + " " + TmpLG("Use") + " " + "�������=" + Str(TmpLG.AbsolutePosition)

'
'if tmplg("")
'
End If

Next j
Next I

TmpLG.Close

'**********************************************************************************
'***********************�������� ��������� ������ ��� �����������******************
'**********************************************************************************

' ������ ������������ ����������� ����� �������
' If TmpLG("Use")="�� ������" or TmpLG("Use")="��/���.���" - �� ����������� ��������� �� ����� ��������
' ������� ����������� If TmpLG("Use")="��� �������" or TmpLG("Use")="�� �� ����" or TmpLG("Use")="�� �� 1-��"
' ����� "�������" TmpLG("Use")="�� ������" or TmpLG("Use")="��/���.���"



' ��������� TmpLG ��� ����������� ����� �.�. TmpLG("Prim")<>-1
' �� ����������� ������ �.�. ��� ����������, ��� ��� �� ����������

TmpLG.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.Plo, tmp_lgota.Prop, tmp_lgota.Cocmin, tmp_lgota.OtheCode, tmp_lgota.parametr, tmp_lgota.itog, tmp_lgota.tarif, tmp_lgota.Itog1, tmp_lgota.Prim, tmp_lgota.PloLG, tmp_lgota.Key, tmp_lgota.Prim1, tmp_lgota.Koll, tmp_lgota.SovmPloLG, Adding.KodKat, Adding.dop FROM Adding RIGHT JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd WHERE (((tmp_lgota.UniKOd)=" + ����� + ") AND ((tmp_lgota.Prim)>=0))"), Mconn, adOpenKeyset, adLockPessimistic

LastIn = 0

TmpLG.MoveFirst
Do While Not TmpLG.EOF
' ���� ���� ������ �� ���� �� TmpLG("Prim") = 1 ����� ����������� ��� ���������
If TmpLG("Use") = "�� ����" Then
'If TmpLG("Use") = "�� ����" Or TmpLG("Use") = "�� �� ����" Then
TmpLG("Prim") = 1
TmpLG("itog1") = 1 - TmpLG("Procent") / 100
TmpLG.UpdateBatch
LastIn = TmpLG.AbsolutePosition

'����������� ��� ���������
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

' ���� ��������� TmpLG("Use") = "�� 1-��"
TmpLG.Requery

KolL = 0
PSumm = 0
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "�� 1-��" Then
'Itog - ��� ������� ������ ������ �� ����� ������� ������
'Itog1 - ��� ����� ������� ������ �� ���� ������� ���������

'������������ ���������� ����� ����� KolL
'� ��������� ������� �� ������ PSumm ��� ������������ ������� ������ % �� �/��
 
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



'������ ����������� ����� ������� ����� �� 1-�� ��� ��� ����� � ���������� � Itog1
TmpLG.Requery
ProcentLMAX = 0
TmpLG.MoveFirst
Do While Not TmpLG.EOF

If TmpLG("Use") = "�� 1-��" Then
P = (100 - TmpLG("Procent")) / 100 '
If TmpLG("Prop") <> 0 Then ProcentLMAX = (PSumm + (TmpLG("Prop") - KolL)) / TmpLG("Prop")

TmpLG("itog1") = ProcentLMAX
TmpLG.UpdateBatch
End If

TmpLG.MoveNext
Loop


'**************************** ������ �� �������***********************
TmpLG.Requery

'
TmpLG.MoveFirst

'***************"��� �������"**************

Do While Not TmpLG.EOF

If TmpLG("plo") <> 0 Then

If TmpLG("Use") = "��� �������" Then

TmpLG("SovmPloLG") = TmpLG("plo") * (100 - TmpLG("Procent")) / 100
TmpLG("Itog") = (100 - TmpLG("Procent")) / 100

If TmpLG("Procent") = 100 Then TmpLG("SovmPloLG") = TmpLG("plo")


TmpLG.UpdateBatch
End If
End If

TmpLG.MoveNext
Loop

'***************"�� �� ����"**************
TmpLG.MoveFirst
Do While Not TmpLG.EOF

If TmpLG("plo") <> 0 Then

If TmpLG("Use") = "�� �� ����" Then
TmpLG("SovmPloLG") = TmpLG("plolg") * (100 - TmpLG("Procent")) / 100
TmpLG("Itog") = (TmpLG("plolg") * (100 - TmpLG("Procent")) / 100) / TmpLG("plo")
TmpLG.UpdateBatch
End If
End If

TmpLG.MoveNext
Loop


'***************"��/���.���" "�� ������" "�� �� 1-��" "�� �� 1-��"**************
SumLGPl = 0

TmpLG.MoveFirst
Ostat = TmpLG("plo")

Do While Not TmpLG.EOF

If TmpLG("plo") <> 0 Then

If TmpLG("Use") = "��/���.���" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "�� �� 1-��" Then
' �������� or TmpLG("Use") = "�� �� 1-��"
'���� ���� ��������� ��������� "��/���.���" "�� ������" � "�� �� 1-��"

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


'������ ����������� ��������� ������� ��� ���� ��������� ����������� �����
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("plo") <> 0 Then
If TmpLG("Use") = "��/���.���" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "�� �� 1-��" Then
' �������� or TmpLG("Use") = "�� �� 1-��"
'���� ���� ��������� ��������� "��/���.���" "�� ������" � "�� �� 1-��"

TmpLG("SovmPloLG") = PloSumm
TmpLG.UpdateBatch
End If
End If
TmpLG.MoveNext
Loop


' ������ �������� ������ ������, ��� �� ������, ������� ����� MAX SovmPloLG
' ���� ����� ����� ���� ���������
Max = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If Max < TmpLG("SovmPloLG") Then
Max = TmpLG("SovmPloLG")
End If
TmpLG.MoveNext
Loop

'����������� Prim=1 Prim=-1
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
If TmpLG("Use") = "�� �� ����" Or TmpLG("Use") = "��� �������" Then I = 1
If TmpLG("Use") = "��/���.���" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "�� �� 1-��" Then j = 1
TmpLG.MoveNext
Loop

If I = 1 And j = 1 Then

TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "�� �� ����" Or TmpLG("Use") = "��� �������" Then TmpLG("Prim") = 1
If TmpLG("Use") = "��/���.���" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "�� �� 1-��" Then TmpLG("Prim") = -1
TmpLG.UpdateBatch
TmpLG.MoveNext
Loop
End If

' ������ ����������� ����� ������� � ����������� � Adding
TmpLG.Requery
TmpLG.MoveFirst
strLg = " "
Do While Not TmpLG.EOF
strLg = strLg + Str(TmpLG("KodKls")) + ","
If TmpLG("Use") = "�� �� ����" Or TmpLG("Use") = "��� �������" Or TmpLG("Use") = "��/���.���" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "�� �� 1-��" Then



TmpLG("Itog1") = (TmpLG("plo") - TmpLG("SovmPloLG")) / TmpLG("plo")
TmpLG.UpdateBatch
End If
TmpLG.MoveNext
Loop
strLg = Trim(strLg)

Mconn.Execute ("UPDATE Adding INNER JOIN TMP_Lgota ON Adding.Key = TMP_Lgota.UniKOd SET Adding.LgotaP = [TMP_Lgota]![Itog1], Adding.LgotaKod = " + Chr(34) + strLg + Chr(34) + " WHERE (((Adding.Key)=" + ����� + ") AND ((TMP_Lgota.Prim)=1))")

''' ����������� ������ � ��������� ���������� ��� ������������ �������

' 1 ���� ����������� 2 � ����� ������ "��� �������" �� ����������� ��� ����� ������
n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "��� �������" Then
n = n + 1
If n > 1 Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop

'2. ���� ����������� 2 � ����� ������ "�� �� ����" ����������� ����� ������� �� ����������� ��� ����� ������

n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "�� �� ����" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop

'3. ���� ����������� 2 � ����� ������ "�� ������" ����������� ����� ������� �� ����������� ��� ����� ������

n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "�� ������" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop


'4. ���� ����������� 2 � ����� ������ "��/���.���" ����������� ����� ������� �� ����������� ��� ����� ������

n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "��/���.���" Then
n = n + TmpLG("plolg")
If TmpLG("plo") < n Then
TmpLG("prim") = -1
TmpLG("plolg") = 0
TmpLG.UpdateBatch

End If
End If
TmpLG.MoveNext
Loop

' 5. ��������� �� ��������� �� ������������ ������� �����


n = 0
TmpLG.Requery
TmpLG.MoveFirst
Do While Not TmpLG.EOF
If TmpLG("Use") = "��/���.���" Or TmpLG("Use") = "�� ������" Or TmpLG("Use") = "�� �� 1-��" Or TmpLG("Use") = "��� �������" Or TmpLG("Use") = "�� �� ����" Then
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
