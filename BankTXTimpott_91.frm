VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BankTXTimpott_91 
   Caption         =   "Form5"
   ClientHeight    =   6228
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7992
   LinkTopic       =   "Form5"
   ScaleHeight     =   6228
   ScaleWidth      =   7992
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1452
   End
   Begin VB.DirListBox Dir1 
      Height          =   2232
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   2532
   End
   Begin VB.FileListBox File1 
      Height          =   2184
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   4692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������� ����"
      Height          =   492
      Left            =   2040
      TabIndex        =   1
      Top             =   4920
      Width           =   5532
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   3360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2652
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   7332
      _ExtentX        =   12933
      _ExtentY        =   656
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "������� ���� � ����� �����. ��������� ���� �� �����  *.??y"
      Height          =   492
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6492
   End
   Begin VB.Label Label2 
      Caption         =   "\"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6132
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Enabled         =   0   'False
      Height          =   252
      Left            =   6120
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label4 
      Caption         =   " "
      Height          =   492
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   7452
   End
   Begin VB.Label Label5 
      Height          =   252
      Left            =   6840
      TabIndex        =   6
      Top             =   720
      Width           =   492
   End
End
Attribute VB_Name = "BankTXTimpott_91"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
MainMenu.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()


'MsgBox (NEOPOZNAN)

Command2.Font.Bold = True
Command2.FontSize = 12


Command2.Caption = "����� ���� ��������� �����"


Dim rsReestrDoc As ADODB.Recordset



fileN = Trim(Me.Label2.Caption)

Set FSO = CreateObject("Scripting.FileSystemObject")
Set F = FSO.OpenTextFile(fileN)
 
 

 
'��������� ���� ���������
Do While Not F.AtEndOfStream
   Stroka = F.ReadLine
   TestArray = Split(Stroka, ";")
     
 ' ���������� ������ ��� ������ ��������� ������ � "=" � ��������� ��������� �������
     
    For I = 0 To UBound(TestArray)
    If TestArray(I) <> "" Then
    '    LastNonEmpty = LastNonEmpty + 1
'        TestArray(LastNonEmpty) = TestArray(i)
        
If InStr(TestArray(I), "=") = 1 Then ' ������� ��������� ������ �����

Kol = Replace(TestArray(I), "=", "") ' ���������� ����� � �������

'��������� ����������

Me.ProgressBar1.Max = Kol + 10
Me.ProgressBar1.Value = 1
Me.ProgressBar1.Visible = True


Sum = TestArray(1) '����� ����� �������� �������
npp = TestArray(4) '����� ���������� ���������
DPP = Left(TestArray(5), 10) '���� ���������� ���������

'��������� ������ �������� �������
DataPP = CDate(DPP)
KomR = "������� " + Sum + " ���. � �/� " + npp + " �� " + DPP + " ���.��������=" + Kol

'MsgBox (KomR)
End If
      
      
      'MsgBox (TestArray(i))
    End If
Next
Loop

KodN = Trim(Left(Combo1.Text, (InStr(Trim(Combo1.Text), " "))))

'��������� ������ � ������ ����������
Mconn.Execute ("INSERT INTO ReestrDoc ( Data, NachCod, Nach, Coment, Summa, Status, Tip, KodDom, Adres ) SELECT '" + Replace(DPP, "-", "/") + "', " + KodN + ", '����','" + KomR + "' , 0, 0, '������ �����', 0, '��� ������'")


    'Mconn.Execute ("INSERT INTO ReestrDoc ( Data, Coment ) SELECT '" + Replace(DPP, "-", "/") + "', '" + KomR + "'")
F.Close
    
    '������� ��� ������ ���������
    Set rsReestrDoc = New ADODB.Recordset
    rsReestrDoc.Open ("SELECT ReestrDoc.Cod FROM ReestrDoc"), Mconn
    
    rsReestrDoc.MoveFirst
    maxs = rsReestrDoc("Cod")
        Do While Not rsReestrDoc.EOF
        
        If maxs < rsReestrDoc("Cod") Then maxs = rsReestrDoc("Cod")
    
rsReestrDoc.MoveNext
         Loop

rsReestrDoc.Close

    Cod = Str(maxs)
    
    
    
                          '��������� ������ ��������
   '��������� ���� ���������
 Set F = FSO.OpenTextFile(fileN)
                    Do While Not F.AtEndOfStream
                    
RaznesenoOK = False '������ ���� ��� ���� ������ � ����
                    

                    
   Stroka = F.ReadLine
   TestArray = Split(Stroka, ";")
     
 '*****************************************************************************************
     
 ' ���������� ������
     
    'For i = 0 To UBound(TestArray)
    If TestArray(0) <> "" And InStr(TestArray(0), "=") = 0 Then
    
    Dpl = Trim(TestArray(0)) ' ���� �������
    Filial = TestArray(2) '����� ���������
    Ls = Trim(TestArray(5)) ' ������� ����
        '��������� ���� �� ���-�� �������� � ������� ����
        '
            
            
     Me.Label4.Caption = "������� �/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7)
     
     
            
            If Me.Label3.Caption = "BanKN" Then  ' ���� ����� �� 12-�������� ������
            Do While Len(Trim(Ls)) < 12
    Ls = "0" + Ls
            Loop
           
         Else
' ��� �� ���� �� OLDNUM �� ������� ������� ����
            
Do While Left(Ls, 1) = 0 ' ������� ��� ������� ����
            
  If Left(Ls, 1) = 0 Then
  Ls = Right(Ls, Len(Ls) - 1)
  'MsgBox ("��������! ������ ������� ���� �� �.��. " + Ls + ".")
  End If
            Loop
            
           
            
            End If
            
Dim rsMain As ADODB.Recordset
'���� �������� � ����
Set rsMain = New ADODB.Recordset


' ������� ����� ����� ��� �� ��������� � ������� ��������

rsMain.Open ("SELECT MainOccupant.Numer, KLS_PODR.��� AS Dom, MainOccupant.OLDNUM, MainOccupant.BanKN FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���"), Mconn

'RSmain.Open ("SELECT MainOccupant.Numer, KLS_PODR.��� AS Dom, MainOccupant.OLDNUM, MainOccupant.BanKN FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Numer)=" + Ls + "))"), Mconn


rsMain.MoveFirst
Do While Not rsMain.EOF
DoEvents
  If rsMain(Trim(Me.Label3.Caption)) = Ls Then
  LS1 = Trim(Str(rsMain("Numer")))
  Dom = Trim(Str(rsMain("Dom")))
  Ls = ""
  RaznesenoOK = True ' ������ ������ ��������
  End If
   rsMain.MoveNext
   Loop
    
' ****************************************************************
    
    
    ' MsgBox (TestArray(8) + "," + TestArray(9))
    NOtd = TestArray(2) '����� ���������
    nkass = TestArray(3) '����� �������/��/����
    FIO = TestArray(6) '���
    Add = TestArray(7) '�����
    Period = Trim(TestArray(8)) ' ������ ������
    Summ = TestArray(9) '����� ��������
    Summ = Replace(Summ, ",", ".")
    SummKomis = TestArray(11) ' ����� �������� �����
    
    
    
    
    
 Komm = "�� " + Left(Period, 2) + "-" + Right(Period, 2) + ". ��/� �" + Trim(npp) + " ���.�����" + Trim(NOtd) + "������" + Trim(nkass)
 
 ' �������� ������� ������� ������, ���� ����� ���� ����� ������� � ���� �������������� ����������
 'If Len(Period) <> 0 Then
 'Period = "01/" + Left(Period, 2) + "/" + Right(Period, 2) '  ������ ������
 'Else
 'Period = "01/01/01"
 'End If
 Period = "01/01/01"
 
 
 
'��������� ������ �1 � ��������

KodN = Trim(TestArray(8))
Summ = Trim(TestArray(9))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then
'Summ = "0"


  If RaznesenoOK Then
 
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
  
End If
    
    
'��������� ������ �2 � ��������
If Trim(TestArray(10)) = "" Then TestArray(10) = 0

KodN = Trim(TestArray(10))
Summ = Trim(TestArray(11))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
  
  End If
  '��������� ������ �3 � ��������

KodN = Trim(TestArray(12))
Summ = Trim(TestArray(13))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then
    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If

 End If
  
  '��������� ������ �4 � ��������

KodN = Trim(TestArray(14))
Summ = Trim(TestArray(15))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
'��������� ������ �5 � ��������

KodN = Trim(TestArray(16))
Summ = Trim(TestArray(17))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
End If
                                         



 
  '��������� ������ �6 � ��������

KodN = Trim(TestArray(18))
Summ = Trim(TestArray(19))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If


 
  '��������� ������ �7 � ��������

KodN = Trim(TestArray(20))
Summ = Trim(TestArray(21))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If


 
  '��������� ������ �8 � ��������

KodN = Trim(TestArray(22))
Summ = Trim(TestArray(23))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If



 
  '��������� ������ �9 � ��������

KodN = Trim(TestArray(24))
Summ = Trim(TestArray(25))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If



 
  '��������� ������ �10 � ��������

KodN = Trim(TestArray(26))
Summ = Trim(TestArray(27))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If




 
  '��������� ������ �11 � ��������

KodN = Trim(TestArray(28))
Summ = Trim(TestArray(29))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If




 
  '��������� ������ �12 � ��������

KodN = Trim(TestArray(30))
Summ = Trim(TestArray(31))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If



 
  '��������� ������ �13 � ��������

KodN = Trim(TestArray(32))
Summ = Trim(TestArray(33))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 'MsgBox (Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If


 
  '��������� ������ �14 � ��������

KodN = Trim(TestArray(34))
Summ = Trim(TestArray(35))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If

  
  '��������� ������ �15 � ��������

KodN = Trim(TestArray(36))
Summ = Trim(TestArray(37))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
  
   
  '��������� ������ �16 � ��������

KodN = Trim(TestArray(38))
Summ = Trim(TestArray(39))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
 
 
 
  
  '��������� ������ �17 � ��������

KodN = Trim(TestArray(40))
Summ = Trim(TestArray(41))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
 
 
  
  '��������� ������ �18 � ��������

KodN = Trim(TestArray(42))
Summ = Trim(TestArray(43))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
  
   
  '��������� ������ �19 � ��������

KodN = Trim(TestArray(44))
Summ = Trim(TestArray(45))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
  
   
  '��������� ������ �20 � ��������

KodN = Trim(TestArray(46))
Summ = Trim(TestArray(47))

'���� �������� ����� �� ������
If Trim(Summ) <> "" Then

    
                              If Not Trim(TestArray(10)) Then     '���� �������� �� ������
                              If Val(Summ) <> 0 Then
    
 If RaznesenoOK Then
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
 Else
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
 End If
  
                                       End If
                                       End If
 End If
  
   
 '**************************************************************
    
    
    End If
    
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    
    
    
    
    
   
                    Loop
       
F.Close

' ��������� ������ ������� �������

Mconn.Execute ("UPDATE doc INNER JOIN nachisleniy ON doc.KodN = nachisleniy.Kod SET doc.NameN = [nachisleniy]![Naim] WHERE (((doc.NameN) Is Null))")

'Mconn.Execute ("UPDATE nachisleniy INNER JOIN Doc ON nachisleniy.Kod = Doc.KodN SET Doc.NameN = [nachisleniy]![Naim]")



MsgBox ("������ ���������. � ������� ������ �������� �" + Cod)
'ReestrDoc.Enabled = True
'MainMenu.Enabled = True
Me.Hide
ReestrDoc.Show
Unload Me
End Sub

Private Sub Dir1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveEr
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Label2.Caption = File1.Path + "\" + FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")

DriveEr:
If Err.Number = 68 Then MsgBox "��� ����� � ���������, ��� ���� ���������"
End Sub

Private Sub File1_Click()
Label2.Caption = File1.Path + "\" + File1.FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")
End Sub

Private Sub Form_Load()
Dim TestArray(202) As String
Dim LastNonEmpty As Integer
Dim FSO
Dim F
Dim fileN As String ' ���� � ��� ����� ��� ������
Dim Stroka As String ' ����������� ������
Dim DataPP As Date
Dim rsCombo As ADODB.Recordset
Dim RsSet As ADODB.Recordset
Dim RaznesenoOK As Boolean

Dim Ls As String '������� ���� �����
Dim TipLs As String '���� �� �������� ������������� ���. ���� �� ����. Setting
Dim NEOPOZNAN As String '����� �/�� ������������ ����


'���������� ����� ����� ���. ����� ��� ����� � ������
Set RsSet = New ADODB.Recordset
RsSet.Open ("SELECT Settings.BankN, Settings.Neo FROM Settings"), Mconn
TipLs = Trim(RsSet("BankN"))
NEOPOZNAN = Trim(RsSet("Neo"))
Me.Label3.Caption = TipLs
RsSet.Close
Me.Label5.Caption = NEOPOZNAN
LastNonEmpty = -1
File1.Pattern = "*.y??;*.txt"
'��������� ���������
Set rsCombo = New ADODB.Recordset
rsCombo.Open ("SELECT nachisleniy.Kod, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy WHERE (((nachisleniy.Tip)='-'))"), Mconn

rsCombo.MoveFirst
Combo1.Text = (Str(rsCombo("Kod")) + " " + rsCombo("Naim"))
Do While Not rsCombo.EOF
Combo1.AddItem (Str(rsCombo("Kod")) + " " + rsCombo("Naim"))
rsCombo.MoveNext
Loop
End Sub





Private Sub Form_Unload(Cancel As Integer)
Me.Hide
ReestrDoc.Show
End Sub
