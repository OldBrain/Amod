VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form4"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   2292
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' ���� �������� ���������� ��� ������ � World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
Dim TableWord As Word.Table
'*****************************************


'�������� ������

'������ ��� ����� ������
nameRP = "I"
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = False


'*************************************
'// ���� ����� ������� ��������� ��������, �� ����� ����� ���

Set DocWord = WordApp.Documents.Open(App.Path + "\" + nameRP + ".doc")
'���������� ���
DocWord.Activate
'��������� ��������� ��������
nameRP = nameRP + "1.doc"

DocWord.SaveAs (App.Path + "\" + nameRP)
DocWord.Close

'������ ����� ��������� Word-a
'Set WordApp = New Word.Application
' ��������� �������� ���������� ��� ��������� ������
WordApp.Options.CheckSpellingAsYouType = False

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\" + nameRP)

'���������� ���
 DocWord.Activate
'��������� ���������
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(1, 3).Range.Text = "MainForm.NamePr"

'TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = "MainForm.Bank"

'TableWord.Cell(2, 3).Select
TableWord.Cell(2, 3).Range.Text = "MainForm.BIK"

'TableWord.Cell(2, 5).Select
TableWord.Cell(2, 5).Range.Text = "MainForm.KS"

'TableWord.Cell(3, 5).Select
TableWord.Cell(3, 5).Range.Text = "MainForm.RS"

'TableWord.Cell(3, 3).Select
TableWord.Cell(3, 3).Range.Text = "MainForm.INN"

'����
TableWord.Cell(6, 1).Range.Text = "��������� ������ " + "MainForm.Label8" + " �."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'������� ����
'TableWord.Cell(2, 7).Select
TableWord.Cell(2, 7).Range.Text = "RsKvit(BanKN)"
' �����
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = "rsNum(NAIM_KLS)"

' ���
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = "RsKvit(FAM)"

'�������
TableWord.Cell(4, 4).Range.Text = "RsKvit(COMSPACE)"

'���������
TableWord.Cell(4, 6).Range.Text = "RsKvit(NLODGERF)"

'������ �����





TableWord.Cell(4, 8).Range.Text = Format("100", "0.00")

TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")






                            '��������� ��� ��������� �� ������
                            



'���� �� ����������� ������ ��� �����
       
       For i = 11 To 23
        
       
       
                    
                                    '****** ����������� ����������
     '��������� ������ � �������
        DocWord.Tables(1).Rows.Add
      TableWord.Cell(i, 1).Range.Text = "NameN"
      TableWord.Cell(i, 2).Range.Text = "edizm"
    
    ' ����� �����
     TableWord.Cell(i, 3).Range.Text = 10
    '�����
    TableWord.Cell(i, 4).Range.Text = Format(10, "0.00")
    s = 0
      
   TableWord.Cell(i, 5).Range.Text = Format(10, "0.00")
   TableWord.Cell(i, 6).Range.Text = Format(10, "0.00")
   TableWord.Cell(i, 7).Range.Text = Format(10, "0.00")
  
  '��������� ����������� ������������ �����
     TableWord.Cell(i, 8).Range.Text = "�"
     
  '��������� �������� �����
     TableWord.Cell(i, 9).Range.Text = "�"
  
  '������� �� ������ �� �����
    s = 100
  
     TableWord.Cell(i, 10).Range.Text = Format(s, "0.00")
     
     If TableWord.Cell(i, 10).Range.Text = TableWord.Cell(i - 1, 10).Range.Text Then
              
     TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
   End If
     
  
       
                                         Next i
'��������� ����

DocWord.Save
 
DocWord.Close

WordApp.Quit

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
'WordApp.Visible = True


rsNum.MoveNext





Unload Jdite


End Sub

