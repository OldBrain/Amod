VERSION 5.00
Begin VB.Form rep_kvit 
   Caption         =   "Form4"
   ClientHeight    =   5724
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6096
   LinkTopic       =   "Form4"
   ScaleHeight     =   5724
   ScaleWidth      =   6096
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "������"
      Height          =   492
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   5652
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��������� ���������� ���������� ��� ����� ���������"
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   5652
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��������� � ����������� ��� ������ ��������"
      Height          =   492
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   5652
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�� ������ �����"
      Height          =   492
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   5652
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���������� � ����� ��������������(�������)"
      Height          =   492
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   5652
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������ ���������(�������)"
      Height          =   492
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   5652
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������ ����������"
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   5652
   End
   Begin VB.CheckBox Check2 
      Caption         =   "��� ��������� ����������"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2892
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�������� �/�� 12 ������"
      Height          =   372
      Left            =   3480
      TabIndex        =   2
      Top             =   5280
      Value           =   1  'Checked
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "� ������� � ������"
      Height          =   492
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   5652
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   5652
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Enabled         =   0   'False
      Height          =   132
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "rep_kvit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'*************************************** 2
Private Declare Sub PDF417Encode Lib "PDF417Font.dll" _
(ByVal Message As String, ByVal Mode As Integer, ByVal ECLevel As Integer, _
 ByVal Rows As Integer, ByVal Columns As Integer, ByVal TruncatedSymbol As Boolean, _
 ByVal HandleTilde As Boolean)

Private Declare Function PDF417GetRows Lib "PDF417Font.dll" () As Integer
Private Declare Function PDF417GetCols Lib "PDF417Font.dll" () As Integer
Private Declare Function PDF417GetCharAt Lib "PDF417Font.dll" (ByVal RowIndex As Integer, ByVal ColIndex As Integer) As Integer

    Dim RowCount As Integer
    Dim ColCount As Integer
    Dim OneLine As String
    Dim EncodedMsg As String
   Public Exit_Me As Boolean
   '��������� ��������� ��� ������ ���������
'**************1*************************
Option Explicit
 
Private Enum TErrorCorretion
    QualityLow
    QualityMedium
    QualityStandard
    QualityHigh
End Enum
 
Private Declare Sub GenerateBMP _
                Lib "quricol32.dll" _
                Alias "GenerateBMPW" ( _
                ByVal FileName As Long, _
                ByVal Text As Long, _
                ByVal Margin As Long, _
                ByVal Size As Long, _
                ByVal Level As TErrorCorretion)

    
   
   

'******************************
Private Sub Command1_Click()
Dim nameRP  As String
Dim O As Object
Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'Dim WordApp1 As Word.Application ' ��������� ����������
'Dim DocWord1 As Word.Document ' ��������� ���������
'Dim S As Integer
'Dim S1 As Integer

'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim i As Integer
Dim s As Double
'*****************************************


'���� �� ������ �����

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'���������� ��� ����
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' ���� ��� ������ ������� ������

Me.Label1.Caption = fil

LSKvit.Show 1

'���� ������ �� �������
If Exit_Me = True Then Exit Sub



'MsgBox (fil)
'�������� ��������� ��� ��������� ������ � �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'�������� ������
'���� ��������� ������� ��� ������ ����
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")

rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")

'�������� ��������� ��� �����
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'�������� ��������� ��� ��������� ������ � ����������� �� �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'�������� ��������� ��� ��������� ������ � ����������� �� ����������
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'���� �� ������� ������ ����
rsNum.MoveFirst
Do While Not rsNum.EOF




'��������� ��� ��������� ������ � ����������� ������ ��� �����
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.TarifD, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** ������ �� ������� �� ���������� ��� ���������� ����� ���������
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.TarifD, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> '���')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "���������� ���������. �������� ����� ���������."
Jdite.Label1 = rsNum("NAIM_KLS") + " �� � " + RsKvit("kv_num") + " ���.����" + rsNum("Oldnum")


' ������ ����. ������
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' ��� ������ � �������� ������ � ����. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'������ ��� ����� ������
nameRP = "I"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + rsNum("NAIM_KLS") + "_�� �_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'������� ����� �� �������� �����
nameRP = Replace(nameRP, ".", "_")

'������� ���� �� �������� �����
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'���������� ���
 DocWord.Activate

'��������� ���������
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(1, 3).Range.Text = MainForm.NamePr

'TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = MainForm.Bank

'TableWord.Cell(2, 3).Select
TableWord.Cell(2, 3).Range.Text = MainForm.BIK

'TableWord.Cell(2, 5).Select
TableWord.Cell(2, 5).Range.Text = MainForm.KS

'TableWord.Cell(3, 5).Select
TableWord.Cell(3, 5).Range.Text = MainForm.RS

'TableWord.Cell(3, 3).Select
TableWord.Cell(3, 3).Range.Text = MainForm.INN

'����
TableWord.Cell(6, 1).Range.Text = "��������� ������ " + MainForm.Label8 + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If

' �����
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num")


' ���

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'�������
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'���������
TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'������ �����



OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
End If
OplataRS.Close




                            '��������� ��� ��������� �� ������
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'���� �� ����������� ������ ��� �����
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        '���������� ������
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** ����������� ����������
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "-"
    End If
    
    
    ' ����� �����
    ' ���� Parametr="���������" �� ������ ���������
    
    If RsKvitK("Parametr") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' ���� Parametr="������" �� ������ ��������� ����� *
    If RsKvitK("Parametr") = "������" Or RsKvitK("SchetZ") = "���" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' ���� Parametr="�������" ��� "�������" �� ������ ��������� ����� �������
    If (RsKvitK("Parametr") = "�������" Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    '�����
    
    If InStr(1, RsKvitK("NameN"), "����") = 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    '������ ����� �� ����
    
    If InStr(1, RsKvitK("NameN"), "����") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   '���� ��� ����� ����������
        
    'S ��� �������� ���� ����� �� ������
    s = 0
        
        If RsKvitK("SchetZ") = "�����" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
                
     s = s + RsKvitK("SummaI")
        End If
        
    '���� ��� ��� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     '���� ��� ����������� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "-"
     
 ' End If
  
  
  '��������� ����������� ������������ �����
  
     If RsKvitK("norm") <> 0 Then
     TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     Else
     TableWord.Cell(i, 8).Range.Text = "�"
     End If
  
  
  '��������� �������� �����
     If RsKvitK("Sch") = "��" Then
     If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
     If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " �� ���������"
     
     Else
     TableWord.Cell(i, 9).Range.Text = "-"
     End If
  
                                        
  '������� �� ������ �� �����
  
  If RsKvitK("saldok") >= 0 Then
     
    ' If TableWord.Cell(i, 10).Range.Text <> TableWord.Cell(i - 1, 10).Range.Text Then
'      TableWord.Cell(i, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
      
     
              
 '    TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
  '   TableWord.Cell(i - 1, 10).Range.Delete
  '   TableWord.Cell(i - 1, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
          
     'End If
     
  End If
  
  If RsKvitK("saldok") < 0 Then
  
  'If TableWord.Cell(i, 11).Range.Text <> TableWord.Cell(i - 1, 11).Range.Text Then
  
  'TableWord.Cell(i, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
  
  
    
     
   '  TableWord.Cell(i, 11).Merge MergeTo:=TableWord.Cell(i - 1, 10)
'     TableWord.Cell(i, 11).Range.Delete
'     TableWord.Cell(i - 1, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
     
   ' End If
    
  
  End If
  
  
                                              End If
                                              
                                              
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
  
  
           RsKvitK.MoveNext
        
  
  
  
  
  'Tables(1).Rows.Add
        
        
 
        Loop
        
        
                                   '�����
                                   
Set TableWord = DocWord.Tables(2)


'����� ���������
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 OplataRS.Close
 
 
 '�������������/��������� �� ������ �������
OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "������������� �� ������� �������"
If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "��������� �� ������� ������� "
If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

End If

  OplataRS.Close
 
 
 '��������
 
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
End If
OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 '�������������/��������� �� ����� ������� ��� �� � ���� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
 End If
 OplataRS.Close
 
 
 
 ' ����� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 End If
 OplataRS.Close
        
        
        '��������� ��� ��������� �� ������
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "������������ ��������� ������� ���������"


Unload Jdite

MsgBox ("������������ ��������� ������� ���������. ����� ��������� ��������� � " + App.Path + "\izv\")

Unload Reports
MainMenu.Enabled = True
Unload Me



End Sub

        
        
       

Private Sub Command2_Click()  ' ��������� ������ ������

Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'Dim WordApp1 As Word.Application ' ��������� ����������
'Dim DocWord1 As Word.Document ' ��������� ���������
'Dim s As Single
'Dim S1 As Integer

'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim nameRP As String
Dim s As Double
Dim i As Integer

'*****************************************


'���� �� ������ �����

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'���������� ��� ����
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' ���� ��� ������ ������� ������

Me.Label1.Caption = fil

LSKvit.Show 1
'���� ������ �� �������
If Exit_Me = True Then Exit Sub







'MsgBox (fil)
'�������� ��������� ��� ��������� ������ � �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'�������� ������
'���� ��������� ������� ��� ������ ����
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")


'�������� ��������� ��� �����
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'�������� ��������� ��� ��������� ������ � ����������� �� �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'�������� ��������� ��� ��������� ������ � ����������� �� ����������
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'���� �� ������� ������ ����
rsNum.MoveFirst
Do While Not rsNum.EOF




'��������� ��� ��������� ������ � ����������� ������ ��� �����
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** ������ �� ������� �� ���������� ��� ���������� ����� ���������
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.TarifD, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> '���')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "���������� ���������. �������� ����� ���������."
Jdite.Label1 = rsNum("NAIM_KLS") + " �� � " + RsKvit("kv_num") + " ���.����" + rsNum("Oldnum")


' ������ ����. ������
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' ��� ������ � �������� ������ � ����. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'������ ��� ����� ������
nameRP = "ibn"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + rsNum("NAIM_KLS") + "_�� �_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'������� ����� �� �������� �����
nameRP = Replace(nameRP, ".", "_")

'������� ���� �� �������� �����
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'���������� ���
 DocWord.Activate

'��������� ���������
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(1, 3).Range.Text = MainForm.NamePr

'TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = MainForm.Bank

'TableWord.Cell(2, 3).Select
TableWord.Cell(2, 3).Range.Text = MainForm.BIK

'TableWord.Cell(2, 5).Select
TableWord.Cell(2, 5).Range.Text = MainForm.KS

'TableWord.Cell(3, 5).Select
TableWord.Cell(3, 5).Range.Text = MainForm.RS

'TableWord.Cell(3, 3).Select
TableWord.Cell(3, 3).Range.Text = MainForm.INN

'����
TableWord.Cell(6, 1).Range.Text = "��������� ������ " + MainForm.Label8 + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If

' �����
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num")


' ���

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'�������
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'���������
TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'������ ����� � ���� ��������� �� �����



'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close




                            '��������� ��� ��������� �� ������
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'���� �� ����������� ������ ��� �����
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        '���������� ������
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** ����������� ����������
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "X"
    End If
    
    
    ' ����� �����
    ' ���� Parametr="���������" �� ������ ���������
    
    If RsKvitK("Parametr") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' ���� Parametr="������" �� ������ ��������� ����� *
    If RsKvitK("Parametr") = "������" Or RsKvitK("SchetZ") = "���" Then
    TableWord.Cell(i, 3).Range.Text = "X"
    End If
    
    ' ���� Parametr="�������" ��� "�������" �� ������ ��������� ����� �������
    If (RsKvitK("Parametr") = "�������" Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    '�����
    
    If InStr(1, RsKvitK("NameN"), "����") = 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    '������ ����� �� ����
    
    If InStr(1, RsKvitK("NameN"), "����") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   '���� ��� ����� ����������
        
    'S ��� �������� ���� ����� �� ������
    s = 0
        
        If RsKvitK("SchetZ") = "�����" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "X"
                
     s = s + RsKvitK("SummaI")
        End If
        
    '���� ��� ��� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "X"
     s = s + RsKvitK("SummaI")
     End If
     
     '���� ��� ����������� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "X"
     s = s + RsKvitK("SummaI")
     End If
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "X"
     
 ' End If
  
  
  '��������� ����������� ������������ �����
  
     If RsKvitK("norm") <> 0 Then
     TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     Else
     TableWord.Cell(i, 8).Range.Text = "�"
     End If
  
  
  '��������� �������� �����
     If RsKvitK("Sch") = "��" Then
     If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
     If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " �� ���������"
     
     Else
     TableWord.Cell(i, 9).Range.Text = "�"
     End If
  
                                        
  '������� �� ������ �� �����
  
  If RsKvitK("saldok") >= 0 Then
     
    ' If TableWord.Cell(i, 10).Range.Text <> TableWord.Cell(i - 1, 10).Range.Text Then
'      TableWord.Cell(i, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
      
     
              
 '    TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
  '   TableWord.Cell(i - 1, 10).Range.Delete
  '   TableWord.Cell(i - 1, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
          
     'End If
     
  End If
  
  If RsKvitK("saldok") < 0 Then
  
  'If TableWord.Cell(i, 11).Range.Text <> TableWord.Cell(i - 1, 11).Range.Text Then
  
  'TableWord.Cell(i, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
  
  
    
     
   '  TableWord.Cell(i, 11).Merge MergeTo:=TableWord.Cell(i - 1, 10)
'     TableWord.Cell(i, 11).Range.Delete
'     TableWord.Cell(i - 1, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
     
   ' End If
    
  
  End If
  
  
                                              End If
                                              
                                              
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
  
  
           RsKvitK.MoveNext
        
  
  
  
  
  'Tables(1).Rows.Add
        
        
 
        Loop
        
        
                                   '�����
                                   
Set TableWord = DocWord.Tables(2)


'����� ���������
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 OplataRS.Close
 
 
 '�������������/��������� �� ������ ������� � ���� ��������� �� �����
 
'OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


'If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "������������� �� ������� �������"
'If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "��������� �� ������� ������� "
'If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



'TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

'End If

 ' OplataRS.Close
 
 
 '�������� � ���� ��������� �� �����
 
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 '�������������/��������� �� ����� ������� ��� �� � ���� � ������ � ���� ��������� �� �����
'OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
' End If
 'OplataRS.Close
 
 
 
 ' ����� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 End If
 OplataRS.Close
        
        
        '��������� ��� ��������� �� ������
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "������������ ��������� ������� ���������"


Unload Jdite

MsgBox ("������������ ��������� ������� ���������. ����� ��������� ��������� � " + App.Path + "\izv\")

Unload Reports
MainMenu.Enabled = True
Unload Me









End Sub

Private Sub Command3_Click() ' �������

Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'Dim WordApp1 As Word.Application ' ��������� ����������
'Dim DocWord1 As Word.Document ' ��������� ���������
'Dim S As Integer
'Dim S1 As Integer

'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim nameRP As String
Dim s As String
Dim i As Integer

Dim CodeVersion As String
 Dim Name1 As String
 Dim PersonalAcc As String
 Dim BankName As String
 Dim BIC As String
 Dim CorrespAcc As String
 Dim PayeeINN As String
 Dim Category As String
 Dim lastName As String
 Dim firstName As String
 Dim middleName As String
 Dim PersAcc As String
 Dim PayerAddress As String
 Dim J As Integer
 
'*****************************************


'���� �� ������ �����

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'���������� ��� ����
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' ���� ��� ������ ������� ������

Me.Label1.Caption = fil

LSKvit.Show 1
'���� ������ �� �������
If Exit_Me = True Then Exit Sub







'MsgBox (fil)
'�������� ��������� ��� ��������� ������ � �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'�������� ������
'���� ��������� ������� ��� ������ ����
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")


'�������� ��������� ��� �����
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'�������� ��������� ��� ��������� ������ � ����������� �� �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'�������� ��������� ��� ��������� ������ � ����������� �� ����������
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'���� �� ������� ������ ����
rsNum.MoveFirst
Do While Not rsNum.EOF




'��������� ��� ��������� ������ � ����������� ������ ��� �����
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** ������ �� ������� �� ���������� ��� ���������� ����� ���������
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> '���')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "���������� ���������. �������� ����� ���������."
Jdite.Label1 = rsNum("NAIM_KLS") + " �� � " + RsKvit("kv_num") + " ���.����" + rsNum("Oldnum")


' ������ ����. ������
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' ��� ������ � �������� ������ � ����. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'������ ��� ����� ������
nameRP = "ipt"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + rsNum("NAIM_KLS") + "_�� �_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'������� ����� �� �������� �����
nameRP = Replace(nameRP, ".", "_")

'������� ���� �� �������� �����
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'���������� ���
 DocWord.Activate

'��������� ���������
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(2, 2).Range.Text = MainForm.NamePr + ", ���:" + MainForm.INN + ", ����:" + MainForm.Bank + ", ���:" + MainForm.BIK + ", ���.����.:" + MainForm.KS + ", �.����:" + MainForm.RS

'TableWord.Cell(2, 1).Select
'TableWord.Cell(2, 1).Range.Text = MainForm.Bank

'TableWord.Cell(2, 3).Select
'TableWord.Cell(2, 3).Range.Text = MainForm.BIK

'TableWord.Cell(2, 5).Select
'TableWord.Cell(2, 5).Range.Text = MainForm.KS

'TableWord.Cell(3, 5).Select
'TableWord.Cell(3, 5).Range.Text = MainForm.RS

'TableWord.Cell(3, 3).Select
'TableWord.Cell(3, 3).Range.Text = MainForm.INN

'����
TableWord.Cell(5, 1).Range.Text = "��������� ������ " + MainForm.Label8 + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(1, 1).Range.Text = "�.����: " + RsKvit("BanKN")
Else
TableWord.Cell(1, 1).Range.Text = "�.����: " + RsKvit("oldnum")
End If

' �����

TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = "�����:" + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + ", �������:" + Str(RsKvit("COMSPACE")) + "��.�., ��������� ���.���.:" + Str(RsKvit("NLODGERF"))


'------*********------��������� ��������-------------------------------
'Dim s As String
 
' �������� ����� ���������
CodeVersion = "ST00011|" ' ����� ��������� CodePage =1 (WIN1251) � ����������� |
Name1 = "Name =" + Replace(MainForm.NamePr, Chr$(34), "'") + "|" '������������ ���������� �������� ����� �������� ������� �� ���������
PersonalAcc = "PersonalAcc=" + MainForm.RS + "|" '����� ����� ���������� ��������
BankName = "BankName =" + MainForm.Bank + "|" '������������ ����� ���������� ��������
BIC = "BIC = " + MainForm.BIK + "|" ' ���� ��� ���
CorrespAcc = "CorrespAcc =" + MainForm.KS + "|" ' �������
PayeeINN = "PayeeINN =" + MainForm.INN + "|" ' ���
Category = "Category =|" ' �������������� ���� ����� �������� ������
lastName = "lastName =" + RsKvit("FAM") + "|" '�������
firstName = "firstName =" + RsKvit("IM") + "|" '���
middleName = "middleName =" + RsKvit("IM") + "|" ' ��������

' ����� �������� �����

If Me.Check1 Then
PersAcc = "PersAcc=" + RsKvit("BanKN") + "|"
Else
PersAcc = "PersAcc=" + RsKvit("oldnum") + "|"
End If

'����� ����� ��������� ������ ����������� �� ���������
PayerAddress = "PayerAddress=" + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num")

' ��������� �������� ���������
s = CodeVersion + Name1 + PersonalAcc + BankName + BIC + CorrespAcc + PayeeINN + Category + lastName + firstName + middleName + PersAcc + PayerAddress




                      ' �� �������� �� Win10 ������ �� XP ������� ������

' ���������� � ������� pdf417
'Set O = CreateObject("pdf417.clspdf417")
'b = O.pdf417(s, -1)
'������� � ���������
'TableWord.Cell(4, 2).Range.Text = b
'Set O = Nothing

'����� ��� *****************

'txtPDF417.Text = ""
'txtPDF417.FontName = MW6PDF417R6.TTF
 'txtPDF417.FontName = cbxFontName.Text
 ' txtPDF417.FontSize = CInt(cbxFontSize.Text)
    
    ' encode string using PDF417
    
    
   ' �������� ������� ������������ ���������
   ' S-��� ������ ������� ��������
   ' 10,10 - ������� ������� ������ �� ������ ���������
    
    
    Call PDF417Encode(s, 2, _
                      2, 10, _
                      10, False, False)
    
    ' how many rows?
    RowCount = PDF417GetRows
    ' how many characters in one row?
    ColCount = PDF417GetCols
    
   
   ' PDF417GetCharAt ���������� ����� ��� �������
    EncodedMsg = vbCrLf
    For i = 1 To RowCount
        For J = 1 To ColCount
            EncodedMsg = EncodedMsg & Chr(PDF417GetCharAt(i - 1, J - 1))
            'MsgBox (EncodedMsg)
        Next J
        EncodedMsg = EncodedMsg & vbCrLf
    Next i
   

'������� � ���������
TableWord.Cell(4, 2).Range.Text = EncodedMsg

'*************************



'------***********---------------------------------------

' ���

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")

End If

'�������
'TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'���������
'TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'������ ����� � ���� ��������� �� �����



'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close




                            '��������� ��� ��������� �� ������
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'���� �� ����������� ������ ��� �����
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        '���������� ������
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** ����������� ����������
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = " "
    End If
    
    
    ' ����� �����
    ' ���� Parametr="���������" �� ������ ���������
    
    If RsKvitK("Parametr") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' ���� Parametr="������" �� ������ ��������� ����� *
    If RsKvitK("Parametr") = "������" Or RsKvitK("SchetZ") = "���" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' ���� Parametr="�������" ��� "�������" �� ������ ��������� ����� �������
    If (RsKvitK("Parametr") = "�������" Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    '�����
    
    If InStr(1, RsKvitK("NameN"), "����") = 0 Then
    
    ' ����� ������ ������ ������ ����� ��������� �������� �������
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "-"
    End If
    
    '������ ����� �� ����
    
    If InStr(1, RsKvitK("NameN"), "����") <> 0 Then
    
    ' ����� ������ ������ ������ ����� ��������� �������� �������
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    
    
    
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   '���� ��� ����� ����������
        
    'S ��� �������� ���� ����� �� ������
    s = 0
        
        If RsKvitK("SchetZ") = "�����" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
                
     s = s + RsKvitK("SummaI")
        End If
        
    '���� ��� ��� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     '���� ��� ����������� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     
     If s <> 0 Then TableWord.Cell(i, 7).Range.Text = Format(s, "0.00") Else TableWord.Cell(i, 7).Range.Text = " "
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = " "
     
 ' End If
  
  
  '��������� ����������� ������������ ����� � ���� ��������� �� �����
  
     'If RsKvitK("norm") <> 0 Then
     'TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     'Else
     'TableWord.Cell(i, 8).Range.Text = "�"
     'End If
  
  
  '��������� �������� ����� � ���� ��������� �� �����
   '  If RsKvitK("Sch") = "��" Then
   '  If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
   '  If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " �� ���������"
     
   '  Else
   '  TableWord.Cell(i, 9).Range.Text = "�"
   '  End If
  
                                        
  '������� �� ������ �� ����� � ���� ��������� �� �����
  
  'If RsKvitK("saldok") >= 0 Then
     
    ' If TableWord.Cell(i, 10).Range.Text <> TableWord.Cell(i - 1, 10).Range.Text Then
'      TableWord.Cell(i, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
      
     
              
 '    TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
  '   TableWord.Cell(i - 1, 10).Range.Delete
  '   TableWord.Cell(i - 1, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
          
     'End If
     
  ''End If
  
  'If RsKvitK("saldok") < 0 Then
  
  'If TableWord.Cell(i, 11).Range.Text <> TableWord.Cell(i - 1, 11).Range.Text Then
  
  'TableWord.Cell(i, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
  
  
    
     
   '  TableWord.Cell(i, 11).Merge MergeTo:=TableWord.Cell(i - 1, 10)
'     TableWord.Cell(i, 11).Range.Delete
'     TableWord.Cell(i - 1, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
     
   ' End If
    
  
  'End If
  
  
                                              End If
                                              
                                              
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
  
  
           RsKvitK.MoveNext
        
  
  
  
  
  'Tables(1).Rows.Add
        
        
 
        Loop
        
        
                                   '�����
                                   
Set TableWord = DocWord.Tables(2)


'����� ���������
 
' DocWord.Tables(2).Rows.Add
 'TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 'OplataRS.Close
 
 
 '�������������/��������� �� ������ ������� � ���� ��������� �� �����
 
'OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


'If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "������������� �� ������� �������"
'If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "��������� �� ������� ������� "
'If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



'TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

'End If

 ' OplataRS.Close
 
 
 '�������� � ���� ��������� �� �����
 
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 '�������������/��������� �� ����� ������� ��� �� � ���� � ������ � ���� ��������� �� �����
'OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
' End If
 'OplataRS.Close
 
 
 
 ' ����� � ������
OplataRS.Open ("SELECT Adding.KodKv, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.Tip From Adding GROUP BY Adding.KodKv, Adding.Tip HAVING (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.Tip)='+'))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 End If
 OplataRS.Close
        
        
        '��������� ��� ��������� �� ������
                 
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext
Loop


Jdite.Label1.Caption = "������������ ��������� ������� ���������"


Unload Jdite

MsgBox ("������������ ��������� ������� ���������. ����� ��������� ��������� � " + App.Path + "\izv\")

Unload Reports
MainMenu.Enabled = True
Unload Me



End Sub

Private Sub Command4_Click()
Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'Dim WordApp1 As Word.Application ' ��������� ����������
'Dim DocWord1 As Word.Document ' ��������� ���������
'Dim S As Integer
'Dim S1 As Integer

'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim nameRP As String
Dim s As Double
Dim i As Integer

'*****************************************


'���� �� ������ �����

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'���������� ��� ����
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' ���� ��� ������ ������� ������

Me.Label1.Caption = fil

LSKvit.Show 1
'���� ������ �� �������
If Exit_Me = True Then Exit Sub







'MsgBox (fil)
'�������� ��������� ��� ��������� ������ � �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'�������� ������
'���� ��������� ������� ��� ������ ����
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")


'�������� ��������� ��� �����
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'�������� ��������� ��� ��������� ������ � ����������� �� �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'�������� ��������� ��� ��������� ������ � ����������� �� ����������
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'���� �� ������� ������ ����
rsNum.MoveFirst
Do While Not rsNum.EOF




'��������� ��� ��������� ������ � ����������� ������ ��� �����
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** ������ �� ������� �� ���������� ��� ���������� ����� ���������
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> '���')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "���������� ���������. �������� ����� ���������."
Jdite.Label1 = rsNum("NAIM_KLS") + " �� � " + RsKvit("kv_num") + " ���.����" + rsNum("Oldnum")


' ������ ����. ������
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' ��� ������ � �������� ������ � ����. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'������ ��� ����� ������
nameRP = "Ipt_z"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + rsNum("NAIM_KLS") + "_�� �_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'������� ����� �� �������� �����
nameRP = Replace(nameRP, ".", "_")

'������� ���� �� �������� �����
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'���������� ���
 DocWord.Activate

'��������� ���������
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(2, 2).Range.Text = MainForm.NamePr + ", ���:" + MainForm.INN + ", ����:" + MainForm.Bank + ", ���:" + MainForm.BIK + ", ���.����.:" + MainForm.KS + ", �.����:" + MainForm.RS

'TableWord.Cell(2, 1).Select
'TableWord.Cell(2, 1).Range.Text = MainForm.Bank

'TableWord.Cell(2, 3).Select
'TableWord.Cell(2, 3).Range.Text = MainForm.BIK

'TableWord.Cell(2, 5).Select
'TableWord.Cell(2, 5).Range.Text = MainForm.KS

'TableWord.Cell(3, 5).Select
'TableWord.Cell(3, 5).Range.Text = MainForm.RS

'TableWord.Cell(3, 3).Select
'TableWord.Cell(3, 3).Range.Text = MainForm.INN

'����
TableWord.Cell(5, 1).Range.Text = "��������� ������ " + MainForm.Label8 + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(1, 1).Range.Text = "�.����: " + RsKvit("BanKN")
Else
TableWord.Cell(1, 1).Range.Text = "�.����: " + RsKvit("oldnum")
End If

' �����

TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = "�����:" + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + ", �������:" + Str(RsKvit("COMSPACE")) + "��.�., ��������� ���.���.:" + Str(RsKvit("NLODGERF"))

' ���

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'�������
'TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'���������
'TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'������ ����� � ���� ��������� �� �����



'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close




                            '��������� ��� ��������� �� ������
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'���� �� ����������� ������ ��� �����
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        '���������� ������
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** ����������� ����������
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = " "
    End If
    
    
    ' ����� �����
    ' ���� Parametr="���������" �� ������ ���������
    
    If RsKvitK("Parametr") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' ���� Parametr="������" �� ������ ��������� ����� *
    If RsKvitK("Parametr") = "������" Or RsKvitK("SchetZ") = "���" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' ���� Parametr="�������" ��� "�������" �� ������ ��������� ����� �������
    If (RsKvitK("Parametr") = "�������" Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    '�����
    
    If InStr(1, RsKvitK("NameN"), "����") = 0 Then
    
    ' ����� ������ ������ ������ ����� ��������� �������� �������
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "-"
    End If
    
    '������ ����� �� ����
    
    If InStr(1, RsKvitK("NameN"), "����") <> 0 Then
    
    ' ����� ������ ������ ������ ����� ��������� �������� �������
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    
    
    
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   '���� ��� ����� ����������
        
    'S ��� �������� ���� ����� �� ������
    s = 0
        
        If RsKvitK("SchetZ") = "�����" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
                
     s = s + RsKvitK("SummaI")
        End If
        
    '���� ��� ��� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     '���� ��� ����������� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     
     If s <> 0 Then TableWord.Cell(i, 7).Range.Text = Format(s, "0.00") Else TableWord.Cell(i, 7).Range.Text = " "
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = " "
     
 ' End If
  
  
  '��������� ����������� ������������ ����� � ���� ��������� �� �����
  
     'If RsKvitK("norm") <> 0 Then
     'TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     'Else
     'TableWord.Cell(i, 8).Range.Text = "�"
     'End If
  
  
  '��������� �������� ����� � ���� ��������� �� �����
   '  If RsKvitK("Sch") = "��" Then
   '  If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
   '  If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " �� ���������"
     
   '  Else
   '  TableWord.Cell(i, 9).Range.Text = "�"
   '  End If
  
                                        
  '������� �� ������ �� ����� � ���� ��������� �� �����
  
  'If RsKvitK("saldok") >= 0 Then
     
    ' If TableWord.Cell(i, 10).Range.Text <> TableWord.Cell(i - 1, 10).Range.Text Then
'      TableWord.Cell(i, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
      
     
              
 '    TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
  '   TableWord.Cell(i - 1, 10).Range.Delete
  '   TableWord.Cell(i - 1, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
          
     'End If
     
  ''End If
  
  'If RsKvitK("saldok") < 0 Then
  
  'If TableWord.Cell(i, 11).Range.Text <> TableWord.Cell(i - 1, 11).Range.Text Then
  
  'TableWord.Cell(i, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
  
  
    
     
   '  TableWord.Cell(i, 11).Merge MergeTo:=TableWord.Cell(i - 1, 10)
'     TableWord.Cell(i, 11).Range.Delete
'     TableWord.Cell(i - 1, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
     
   ' End If
    
  
  'End If
  
  
                                              End If
                                              
                                              
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
  
  
           RsKvitK.MoveNext
        
  
  
  
  
  'Tables(1).Rows.Add
        
        
 
        Loop
        
        
                                   '�����
                                   
Set TableWord = DocWord.Tables(2)


'����� ���������
 
' DocWord.Tables(2).Rows.Add
 'TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 'OplataRS.Close
 
 
 '�������������/��������� �� ������ ������� � ���� ��������� �� �����
 
'OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


'If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "������������� �� ������� �������"
'If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "��������� �� ������� ������� "
'If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



'TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

'End If

 ' OplataRS.Close
 
 
 '�������� � ���� ��������� �� �����
 
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 '�������������/��������� �� ����� ������� ��� �� � ���� � ������ � ���� ��������� �� �����
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


 TableWord.Cell(2, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")

'TableWord.Cell(2, 2).Range.Text = "�� �� ��"

 End If
OplataRS.Close
 
 
 
 ' ����� � ������
OplataRS.Open ("SELECT Adding.KodKv, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.Tip From Adding GROUP BY Adding.KodKv, Adding.Tip HAVING (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.Tip)='+'))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(1, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 End If
 OplataRS.Close
        
        
        '��������� ��� ��������� �� ������
                 
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext
Loop


Jdite.Label1.Caption = "������������ ��������� ������� ���������"


Unload Jdite

MsgBox ("������������ ��������� ������� ���������. ����� ��������� ��������� � " + App.Path + "\izv\")

Unload Reports
MainMenu.Enabled = True
Unload Me




End Sub

Private Sub Command5_Click()

Dim O As Object
Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'Dim WordApp1 As Word.Application ' ��������� ����������
'Dim DocWord1 As Word.Document ' ��������� ���������
'Dim S As Integer
'Dim S1 As Integer

'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim Spravka As String
Dim Pusto As String ' ��� ������ ���������
Dim nameRP As String
Dim SpravkaO As String
Dim i As Integer
Dim s As Double
Dim SpravkaN As String
Dim SpravkaZP As String
Dim SpravkaD As String
Dim strQR As String
Dim strQR1 As String
Dim CodeVersion As String
 Dim Name1 As String
 Dim PersonalAcc As String
 Dim BankName As String
 Dim BIC As String
 Dim CorrespAcc As String
 Dim PayeeINN As String
 Dim Category As String
 Dim lastName As String
 Dim firstName As String
 Dim middleName As String
 Dim PersAcc As String
 Dim PayerAddress As String
 Dim SumQR As String
 
'*****************************************


'���� �� ������ �����

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If

'���������� ��� ����
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' ���� ��� ������ ������� ������

Me.Label1.Caption = fil

LSKvit.Show 1

'���� ������ �� �������
If Exit_Me = True Then Exit Sub

'MsgBox (fil)
'�������� ��������� ��� ��������� ������ � �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'�������� ������
'���� ��������� ������� ��� ������ ����
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")

rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")

'�������� ��������� ��� �����
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs, Settings.Kvit FROM Settings")

'�������� ��������� ��� ��������� ������ � ����������� �� �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'�������� ��������� ��� ��������� ������ � ����������� �� ����������
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn




                                '��������� ��� ��������� �� ������
                            'If rsNum.EOF = False Or rsNum.BOF = False Then

'���� �� ������� ������ ����
rsNum.MoveFirst
Do While Not rsNum.EOF




'��������� ��� ��������� ������ � ����������� ������ ��� �����
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** ������ �� ������� �� ���������� ��� ���������� ����� ���������
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.TarifD, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.LgotaVid, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> '���')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "���������� ���������. �������� ����� ���������."
Jdite.Label1 = rsNum("NAIM_KLS") + " �� � " + rsNum("kv_num") + " ���.����" + rsNum("Oldnum")


' ������ ����. ������
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' ��� ������ � �������� ������ � ����. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'������ ��� ����� ������
nameRP = "lift"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + rsNum("NAIM_KLS") + "_�� �_" + rsNum("kv_num") + "_" + rsNum("Oldnum")

'������� ����� �� �������� �����
nameRP = Replace(nameRP, ".", "_")

'������� ���� �� �������� �����
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

' ������� * �� �������� �����
nameRP = Replace(nameRP, "*", "")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'���������� ���
 DocWord.Activate

'��������� ���������
Set TableWord = DocWord.Tables(1)

'��������� ���������� ����������
TableWord.Cell(1, 1).Range.Text = RsRec("Kvit")

'TableWord.Cell(1, 3).Select
TableWord.Cell(1, 3).Range.Text = MainForm.NamePr

'TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = MainForm.Bank

'TableWord.Cell(2, 3).Select
TableWord.Cell(2, 3).Range.Text = MainForm.BIK

'TableWord.Cell(2, 5).Select
TableWord.Cell(2, 5).Range.Text = MainForm.KS

'TableWord.Cell(3, 5).Select
TableWord.Cell(3, 5).Range.Text = MainForm.RS

'TableWord.Cell(3, 3).Select
TableWord.Cell(3, 3).Range.Text = MainForm.INN

'����
TableWord.Cell(6, 1).Range.Text = "��������� ������ " + MainForm.Label8 + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select
' ��� ������������� ������ ������ OLDNUM �������
Me.Check1 = False



                                   '��������� ��� ��������� �� ������
                            If RsKvit.EOF = False Or RsKvit.BOF = False Then

'If RsKvit("oldnum") = "" Then MsgBox ("")
If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If


' �����
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " �� �" + rsNum("kv_num")





'------*********------��������� ��������-------------------------------

 
 
 
 
' �������� ����� ���������
CodeVersion = "ST00012|" ' ����� ��������� CodePage =1 (WIN1251) � ����������� |
Name1 = "Name=" + Replace(MainForm.NamePr, Chr$(34), "'") + "|" '������������ ���������� �������� ����� �������� ������� �� ���������
PersonalAcc = "PersonalAcc=" + MainForm.RS + "|" '����� ����� ���������� ��������
BankName = "BankName=" + MainForm.Bank + "|" '������������ ����� ���������� ��������
BIC = "BIC=" + MainForm.BIK + "|" ' ���� ��� ���
CorrespAcc = "CorrespAcc=" + MainForm.KS + "|" ' �������
PayeeINN = "PayeeINN=" + MainForm.INN + "|" ' ���
Category = "Category=|" ' �������������� ���� ����� �������� ������
lastName = "lastName=" + RsKvit("FAM") + "|" '�������
firstName = "firstName=" + RsKvit("IM") + "|" '���
middleName = "middleName=" + RsKvit("OT") + "|" ' ��������

' ����� �������� �����

If Me.Check1 Then
PersAcc = "PersAcc=" + RsKvit("BanKN")
Else
PersAcc = "PersAcc=" + RsKvit("oldnum")
End If

'����� ����� ��������� ������ ����������� �� ���������
PayerAddress = "PayerAddress=" + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + "|"

' ��������� �������� ���������
strQR = CodeVersion + Name1 + PersonalAcc + BankName + BIC + CorrespAcc
strQR = Trim(strQR)

strQR1 = PayeeINN + lastName + firstName + middleName + PayerAddress + PersAcc
strQR1 = Trim(strQR1)


' ���


If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'�������
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'���������
TableWord.Cell(4, 6).Range.Text = Str(RsKvit("NLODGER")) + "/" + Str(RsKvit("NLODGER"))

'������ �����


OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
SpravkaO = "�������� � ������� �������:" + Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
Spravka = ""
End If
OplataRS.Close



                                 Else ' ����� ������� �������������� � ������ ���������
                                Pusto = Pusto + rsNum("NAIM_KLS") + " �� �" + rsNum("kv_num") + Chr(13) + Chr(10)


                               End If





                            '��������� ��� ��������� �� ������
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'���� �� ����������� ������ ��� �����
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        '���������� ������
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** ����������� ����������
                           
                           
                           If RsKvitK("Tip") = "+" And RsKvitK("SummaI") <> 0 Then
  
  
   '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "-"
    End If
    
    
    ' ����� �����
    ' ���� LgotaVid="���������" �� ������ ���������
    
    If RsKvitK("LgotaVid") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' ���� Parametr="������" �� ������ ��������� ����� *
    If RsKvitK("LgotaVid") = "������" Or RsKvitK("SchetZ") = "���" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' ���� Parametr="�������" ��� "�������" �� ������ ��������� ����� �������
    If (RsKvitK("LgotaVid") = "����� ��." Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    '�����
    
    '���� ���� �� �����������
    If RsKvitK("LgotaVid") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("TarifD"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    '���� ���� �� �������
    
    If RsKvitK("SchetZ") <> "���" Then
    
    If (RsKvitK("LgotaVid") = "����� ��." Or RsKvitK("LgotaVid") = "���. ��.") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("Tarif")
    If (RsKvitK("LgotaVid") = "���������" Or RsKvitK("LgotaVid") = "���������") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
    
    End If
    
    'If (RsKvitK("LgotaVid") = "����� ��." Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    'TableWord.Cell(I, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    'End If
    
    
      
   '���� ��� ����� ����������
        
    'S ��� �������� ���� ����� �� ������
    s = 0
        
        If RsKvitK("SchetZ") = "�����" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
            
     s = s + RsKvitK("SummaI")
        End If
        
    '���� ��� ��� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     '���� ��� ����������� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     'MsgBox (RsKvitK("SummaI") + "  " + Format(s, "0.00"))
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "-"
     
 
                                        
  '������� �� ������ �� �����
  
  If RsKvitK("saldok") >= 0 Then
     
    ' If TableWord.Cell(i, 10).Range.Text <> TableWord.Cell(i - 1, 10).Range.Text Then
'      TableWord.Cell(i, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
      
     
              
 '    TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
  '   TableWord.Cell(i - 1, 10).Range.Delete
  '   TableWord.Cell(i - 1, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
          
     'End If
     
  End If
  
  If RsKvitK("saldok") < 0 Then
  
  'If TableWord.Cell(i, 11).Range.Text <> TableWord.Cell(i - 1, 11).Range.Text Then
  
  'TableWord.Cell(i, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
  
  
    
     
   '  TableWord.Cell(i, 11).Merge MergeTo:=TableWord.Cell(i - 1, 10)
'     TableWord.Cell(i, 11).Range.Delete
'     TableWord.Cell(i - 1, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
     
   ' End If
    
  
  End If
  
  
                                              End If
                                              
                                              
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
  
  
           RsKvitK.MoveNext
        
  
  
  
  
  'Tables(1).Rows.Add
        
        
 
        Loop
        
        
        
        

        
                                   '�����
                                   
Set TableWord = DocWord.Tables(2)



'����� ���������
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
 
 '��������� ��� ��������� �� ������
'���� ������ � ����� ������� ���� � ���������������
  If OplataRS("Sum-SummaI") = 0 Then
  Pusto = Pusto + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + Chr(13) + Chr(10)
                            End If
 
 
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 3).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 
 ' ����� ��� ���������
 
 
 SumQR = Format(OplataRS("Sum-SummaI"), "0.00")
 SumQR = Replace(SumQR, ".", "")
 SumQR = Replace(SumQR, ",", "")
 SumQR = "Sum=" + SumQR + "|"
 strQR = strQR + SumQR + strQR1
 
 'MsgBox (strQR)
 
 'OplataRS ("Sum-SummaI")
 SpravkaN = "����� ��������� � ������� �������: " + Str(OplataRS("Sum-SummaI"))
 
 Else
 '��������� ��� ��������� �� ������
'���� ������ � ����� ������� ���� � ���������������
 
  Pusto = Pusto + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + Chr(13) + Chr(10)
                             
 End If
 OplataRS.Close
 
 
 
 '�������������/��������� �� ������ �������
OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


If (OplataRS("plus") + OplataRS("minus")) > 0 Then SpravkaZP = "������������� �� ������� �������"
If (OplataRS("plus") + OplataRS("minus")) < 0 Then SpravkaZP = "��������� �� ������� ������� "
If (OplataRS("plus") + OplataRS("minus")) = 0 Then SpravkaZP = ""



SpravkaZP = SpravkaZP + ": " + Format((OplataRS("plus") + OplataRS("minus")), "0.00")

End If

  OplataRS.Close
 
 
 '��������
 
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then

SpravkaO = "��������� ������ � ������� �������: " + Format(OplataRS("Sum-SummaI"), "0.00")

'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
SpravkaO = "��������� ������ � ������� �������: " + Format(0, "0.00")

End If
OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 '�������������/��������� �� ����� ������� ��� �� � ���� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
 End If
 OplataRS.Close
 
 
 
 ' ����� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 'TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 
 If OplataRS("Sum-SK") < 0 Then SpravkaD = "����� ��������� �� ����� �������: " + Format(OplataRS("Sum-SK"), "0.00")
 If OplataRS("Sum-SK") >= 0 Then SpravkaD = "����� ���� �� ����� �������: " + Format(OplataRS("Sum-SK"), "0.00")
 
 End If
 OplataRS.Close
        
 
        
        
        '��������� ��� ��������� �� ������
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
'Set TableWord = DocWord.Tables(3)
       
      '���������� ����������
    'TableWord.Cell(2, 1).Range.Text = "���������(��� ������ ��������):" + Chr(13) + Chr(10) + SpravkaZP + Chr(13) + Chr(10) + SpravkaO + Chr(13) + Chr(10) + SpravkaN + Chr(13) + Chr(10) + SpravkaD
    TableWord.Cell(2, 1).Range.Text = "���������(��� ������ ��������): " + "" + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
    
    'TableWord.Cell(1, 1).Range.Text = "���������(��� ������ ��������):" + "; " + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
       
       
       
'��������� �������� QR-Code
       
     
     
     
     
     
     
     
     
     
'"������ �������!" +
       
       
      strQR = Replace(strQR, " ", "")
      
    GenerateBMP StrPtr("C:\Example.bmp"), StrPtr(strQR), 3, 5, QualityLow
    
    
    DocWord.Shapes.AddPicture "C:\Example.bmp", , True, 235, 0, 100, 70
    
    
    
      
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "������������ ��������� ������� ���������"


Unload Jdite

MsgBox ("������������ ��������� ������� ���������. ����� ��������� ��������� � " + App.Path + "\izv\")

If Len(Pusto) <> 0 Then

MsgBox ("���������� ������ ���������" + Chr(13) + Chr(10) + Pusto)

End If


Unload Reports
MainMenu.Enabled = True
Unload Me



End Sub



Private Sub Command6_Click()

Dim O As Object
Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'Dim WordApp1 As Word.Application ' ��������� ����������
'Dim DocWord1 As Word.Document ' ��������� ���������
'Dim S As Integer
'Dim S1 As Integer

'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim Spravka As String
Dim Pusto As String ' ��� ������ ���������
Dim nameRP As String
Dim s As Double
Dim i As Integer
Dim SpravkaO As String
Dim SpravkaN As String
Dim SpravkaZP As String
Dim SpravkaD  As String
'*****************************************


'���� �� ������ �����

If Combo1.Text = "������ �����" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If

'���������� ��� ����
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' ���� ��� ������ ������� ������

Me.Label1.Caption = fil

LSKvit.Show 1

'���� ������ �� �������
If Exit_Me = True Then Exit Sub

'MsgBox (fil)
'�������� ��������� ��� ��������� ������ � �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'�������� ������
'���� ��������� ������� ��� ������ ����
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")

rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.�������, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")

'�������� ��������� ��� �����
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs, Settings.Kvit FROM Settings")

'�������� ��������� ��� ��������� ������ � ����������� �� �����������
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'�������� ��������� ��� ��������� ������ � ����������� �� ����������
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn




                                '��������� ��� ��������� �� ������
                            'If rsNum.EOF = False Or rsNum.BOF = False Then

'���� �� ������� ������ ����
rsNum.MoveFirst
Do While Not rsNum.EOF




'��������� ��� ��������� ������ � ����������� ������ ��� �����
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** ������ �� ������� �� ���������� ��� ���������� ����� ���������
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.TarifD, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.LgotaVid, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.��� = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> '���')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "���������� ���������. �������� ����� ���������."
Jdite.Label1 = rsNum("NAIM_KLS") + " �� � " + rsNum("kv_num") + " ���.����" + rsNum("Oldnum")


' ������ ����. ������
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' ��� ������ � �������� ������ � ����. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'������ ��� ����� ������
nameRP = "smol"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + rsNum("NAIM_KLS") + "_�� �_" + rsNum("kv_num") + "_" + rsNum("Oldnum")

'������� ����� �� �������� �����
nameRP = Replace(nameRP, ".", "_")

'������� ���� �� �������� �����
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'���������� ���
 DocWord.Activate

'��������� ���������
Set TableWord = DocWord.Tables(1)



'��������� ���������� ����������
TableWord.Cell(1, 1).Range.Text = RsRec("Kvit")

'TableWord.Cell(1, 3).Select
TableWord.Cell(1, 3).Range.Text = MainForm.NamePr

'TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = MainForm.Bank

'TableWord.Cell(2, 3).Select
TableWord.Cell(2, 3).Range.Text = MainForm.BIK

'TableWord.Cell(2, 5).Select
TableWord.Cell(2, 5).Range.Text = MainForm.KS

'TableWord.Cell(3, 5).Select
TableWord.Cell(3, 5).Range.Text = MainForm.RS

'TableWord.Cell(3, 3).Select
TableWord.Cell(3, 3).Range.Text = MainForm.INN

'����
TableWord.Cell(6, 1).Range.Text = "��������� ������ " + MainForm.Label8 + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select
' ��� ������������� ������ ������ OLDNUM �������
'Me.Check1 = False



                                   '��������� ��� ��������� �� ������
                            If RsKvit.EOF = False Or RsKvit.BOF = False Then

'If RsKvit("oldnum") = "" Then MsgBox ("")
If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If


' �����
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " �� �" + rsNum("kv_num")


' ���

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'�������
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'���������
TableWord.Cell(4, 6).Range.Text = Str(RsKvit("NLODGER")) + "/" + Str(RsKvit("NLODGER"))

'������ �����


OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
SpravkaO = "�������� � ������� �������:" + Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
Spravka = ""
End If
OplataRS.Close



                                 Else ' ����� ������� �������������� � ������ ���������
                                Pusto = Pusto + rsNum("NAIM_KLS") + " �� �" + rsNum("kv_num") + Chr(13) + Chr(10)


                               End If





                            '��������� ��� ��������� �� ������
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'���� �� ����������� ������ ��� �����
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        '���������� ������
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** ����������� ����������
                           
                           
                           If RsKvitK("Tip") = "+" And RsKvitK("SummaI") <> 0 Then
  
  
   '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "-"
    End If
    
    
    ' ����� �����
    ' ���� LgotaVid="���������" �� ������ ���������
    
    If RsKvitK("LgotaVid") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' ���� Parametr="������" �� ������ ��������� ����� *
    If RsKvitK("LgotaVid") = "������" Or RsKvitK("SchetZ") = "���" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' ���� Parametr="�������" ��� "�������" �� ������ ��������� ����� �������
    If (RsKvitK("LgotaVid") = "����� ��." Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    '�����
    
    '���� ���� �� �����������
    If RsKvitK("LgotaVid") = "���������" And RsKvitK("SchetZ") <> "���" Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("TarifD"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    '���� ���� �� �������
    
    If RsKvitK("SchetZ") <> "���" Then
    
    If (RsKvitK("LgotaVid") = "����� ��." Or RsKvitK("LgotaVid") = "���. ��.") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("Tarif")
    If (RsKvitK("LgotaVid") = "���������" Or RsKvitK("LgotaVid") = "���������") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
    
    End If
    
    'If (RsKvitK("LgotaVid") = "����� ��." Or RsKvitK("Parametr") = "�������") And RsKvitK("SchetZ") <> "���" Then
    'TableWord.Cell(I, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    'End If
    
    
      
   '���� ��� ����� ����������
        
    'S ��� �������� ���� ����� �� ������
    s = 0
        
        If RsKvitK("SchetZ") = "�����" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
            
     s = s + RsKvitK("SummaI")
        End If
        
    '���� ��� ��� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     '���� ��� ����������� ����������"
     If RsKvitK("SchetZ") = "���" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "-"
     
  
                                        
  '������� �� ������ �� �����
  
  If RsKvitK("saldok") >= 0 Then
     
    ' If TableWord.Cell(i, 10).Range.Text <> TableWord.Cell(i - 1, 10).Range.Text Then
'      TableWord.Cell(i, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
      
     
              
 '    TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
  '   TableWord.Cell(i - 1, 10).Range.Delete
  '   TableWord.Cell(i - 1, 10).Range.Text = Format(RsKvitK("saldok"), "0.00")
          
     'End If
     
  End If
  
  If RsKvitK("saldok") < 0 Then
  
  'If TableWord.Cell(i, 11).Range.Text <> TableWord.Cell(i - 1, 11).Range.Text Then
  
  'TableWord.Cell(i, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
  
  
    
     
   '  TableWord.Cell(i, 11).Merge MergeTo:=TableWord.Cell(i - 1, 10)
'     TableWord.Cell(i, 11).Range.Delete
'     TableWord.Cell(i - 1, 11).Range.Text = Format((RsKvitK("saldok") * -1), "0.00")
     
   ' End If
    
  
  End If
  
  
                                              End If
                                              
                                              
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
  
  
           RsKvitK.MoveNext
        
  
  
  
  
  'Tables(1).Rows.Add
        
        
 
        Loop
        
        
        
        

        
                                   '�����
                                   
Set TableWord = DocWord.Tables(2)



'����� ���������
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
 
 '��������� ��� ��������� �� ������
'���� ������ � ����� ������� ���� � ���������������
  If OplataRS("Sum-SummaI") = 0 Then
  Pusto = Pusto + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + Chr(13) + Chr(10)
                            End If
 
 
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 3).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 'OplataRS ("Sum-SummaI")
 SpravkaN = "����� ��������� � ������� �������: " + Str(OplataRS("Sum-SummaI"))
 
 Else
 '��������� ��� ��������� �� ������
'���� ������ � ����� ������� ���� � ���������������
 
  Pusto = Pusto + rsNum("NAIM_KLS") + " �� �" + RsKvit("kv_num") + Chr(13) + Chr(10)
                             
 End If
 OplataRS.Close
 
 
 
 '�������������/��������� �� ������ �������
OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


If (OplataRS("plus") + OplataRS("minus")) > 0 Then SpravkaZP = "������������� �� ������� �������"
If (OplataRS("plus") + OplataRS("minus")) < 0 Then SpravkaZP = "��������� �� ������� ������� "
If (OplataRS("plus") + OplataRS("minus")) = 0 Then SpravkaZP = ""



SpravkaZP = SpravkaZP + ": " + Format((OplataRS("plus") + OplataRS("minus")), "0.00")

End If

  OplataRS.Close
 
 
 '��������
 
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then

SpravkaO = "��������� ������ � ������� �������: " + Format(OplataRS("Sum-SummaI"), "0.00")

'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
SpravkaO = "��������� ������ � ������� �������: " + Format(0, "0.00")

End If
OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 '�������������/��������� �� ����� ������� ��� �� � ���� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
 End If
 OplataRS.Close
 
 
 
 ' ����� � ������
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 'TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 
 If OplataRS("Sum-SK") < 0 Then SpravkaD = "����� ��������� �� ����� �������: " + Format(OplataRS("Sum-SK"), "0.00")
 If OplataRS("Sum-SK") >= 0 Then SpravkaD = "����� ���� �� ����� �������: " + Format(OplataRS("Sum-SK"), "0.00")
 
 End If
 OplataRS.Close
        
 
        
        
        '��������� ��� ��������� �� ������
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
'Set TableWord = DocWord.Tables(3)
       
      '���������� ����������
    'TableWord.Cell(2, 1).Range.Text = "���������(��� ������ ��������):" + Chr(13) + Chr(10) + SpravkaZP + Chr(13) + Chr(10) + SpravkaO + Chr(13) + Chr(10) + SpravkaN + Chr(13) + Chr(10) + SpravkaD
    TableWord.Cell(2, 1).Range.Text = "���������(��� ������ ��������): " + "" + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
    
    'TableWord.Cell(1, 1).Range.Text = "���������(��� ������ ��������):" + "; " + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
       
       
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "������������ ��������� ������� ���������"


Unload Jdite

MsgBox ("������������ ��������� ������� ���������. ����� ��������� ��������� � " + App.Path + "\izv\")

If Len(Pusto) <> 0 Then

MsgBox ("���������� ������ ���������" + Chr(13) + Chr(10) + Pusto)

End If


Unload Reports
MainMenu.Enabled = True
Unload Me



End Sub



Private Sub Command7_Click()
KvitShapka.Show 1
End Sub



Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
Exit_Me = False




'��������� ���������
Dim Addrconn As ADODB.Recordset

Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic

Addrconn.Open ("SELECT KLS_PODR.���, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.�������������, KLS_PODR.���� From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Combo1.Text = "������ �����"


Addrconn.MoveFirst
Combo1.AddItem "��� ����"
Do While Not Addrconn.EOF
If Addrconn("���") <> -1 Then
Combo1.AddItem Trim(Str(Addrconn("���"))) + " " + Addrconn("NAIM_KLS") + " ��� � " + Addrconn("Num")
End If
Addrconn.MoveNext
Loop
End Sub

Private Sub QR(s As String)
Dim O As Object
Dim a As String
'S = "���������� � ������ ������ ��������� ���    �.��.   67,3    8,30    558,59      558,59  �   � ��� ������� �.���   10  17,10   171,00      171,00   8.03(�.���)     767(�.���) ��� ��� �.���   67,3    0,00    19,17       19,17   �   � ���� �������    �.���   10  17,85   178,50      178,50   8.03(�.���)     767(�.���) �������������� �������  ��� 200 3,76    752,00      752,00   90(���)     14595(���) ��� ��������������  ���.    67,3    0,00    59,74       59,74   �   � ������� ���.    X   35,00   35,00       35,00   �   � ����� ������    ���.    4   82,00   328,00      328,00  �   �"
Set O = CreateObject("pdf417.clspdf417")

Debug.Print O.pdf417(s, -1)
Set O = Nothing
End Sub

