VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BankTXTimpott 
   AutoRedraw      =   -1  'True
   Caption         =   "������ ��������� ������ *h "
   ClientHeight    =   5784
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7884
   LinkTopic       =   "Form5"
   ScaleHeight     =   5784
   ScaleWidth      =   7884
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   372
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   7332
      _ExtentX        =   12933
      _ExtentY        =   656
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   3600
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   840
      Width           =   2652
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������� ����"
      Height          =   492
      Left            =   2280
      TabIndex        =   6
      Top             =   5040
      Width           =   5532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   492
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   1452
   End
   Begin VB.FileListBox File1 
      Height          =   2184
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   4692
   End
   Begin VB.DirListBox Dir1 
      Height          =   2232
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2532
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label Label5 
      Height          =   252
      Left            =   7080
      TabIndex        =   11
      Top             =   840
      Width           =   492
   End
   Begin VB.Label Label4 
      Caption         =   " "
      Height          =   492
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   7452
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Enabled         =   0   'False
      Height          =   252
      Left            =   6360
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label2 
      Caption         =   "\"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   6132
   End
   Begin VB.Label Label1 
      Caption         =   "������� ���� � ����� �����. ��������� ���� �� �����  *.??y"
      Height          =   492
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6492
   End
End
Attribute VB_Name = "BankTXTimpott"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Unload Me
MainMenu.Enabled = True
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

'****************************************************
    If Me.Label3.Caption = "BanKN" Then  ' ���� ����� �� 12-�������� ������
  '�� ����������� LSN MainOccupant.BanKN
    Lsn = rsMain("BanKN")
    Else
  '�� ����������� LSN MainOccupant.OLDNUM
    Lsn = rsMain("OLDNUM")
  '������ ������� ������� ����
  Do While Left(Lsn, 1) = 0 ' ������� ��� ������� ����
            
  If Left(Lsn, 1) = 0 Then
  Lsn = Right(Lsn, Len(Lsn) - 1)
  'MsgBox ("��������! ������ ������� ���� �� �.��. " + Ls + ".")
  End If
            Loop
  
  
  
    End If

'****************************************************






  If Lsn = Ls Then
  
  
  
  
  'If rsMain(Trim(Me.Label3.Caption)) = Ls Then
  LS1 = Trim(Str(rsMain("Numer")))
  Dom = Trim(Str(rsMain("Dom")))
  Ls = ""
  RaznesenoOK = True ' ������ ������ ��������
  End If
   rsMain.MoveNext
   Loop
    
' ****************************************************************
    
    FIO = TestArray(6) '���
    Add = TestArray(7) '�����
    NOtd = TestArray(2) '����� ���������
    nkass = TestArray(3) '����� �������/��/����
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
 
 
 If RaznesenoOK Then

 
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + LS1 + " AS ���������3, '" + FIO + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
    
 Else
  
  
If Dom = "" Then Dom = "-1"
  
Komm = "�/��. " + Trim(TestArray(5)) + " ��� " + TestArray(6) + " �����:" + TestArray(7) + "������ �� " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS ���������1, " + KodN + " AS ���������2, " + Me.Label5.Caption + " AS ���������3, '" + "������������ �����" + "' AS ���������4, " + Summ + " AS ���������5, '-' AS ���������6, " + npp + " AS ���������7, '" + Replace(Dpl, "-", "/") + "' AS ���������8, '" + Komm + "' AS ���������9, " + Dom + " AS ���������10, '" + Period + "' AS ���������11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("��������! ������� � ��� ������ " + Ls + "�� ������ � ����. ������� " + FIO + " ����� " + Add + " ����� " + TestArray(9) + ". ������� � ������������ �����.")
  
  
 End If
  
  
    End If
    
    
    
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    
    

    
    
   
                    Loop
       
F.Close

' ��������� ������ ������� �������

Mconn.Execute ("UPDATE doc INNER JOIN nachisleniy ON doc.KodN = nachisleniy.Kod SET doc.NameN = [nachisleniy]![Naim] WHERE (((doc.NameN) Is Null))")

MsgBox ("������ ���������. � ������� ������ �������� �" + Cod)
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
