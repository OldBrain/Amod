Attribute VB_Name = "WordRep"
Option Explicit

Public Sub �����(nameRP As String)

Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long

'nameRP = "lc.doc"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP)

'���������� ���
DocWord.Activate

'��������� ��������� ��������
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)
'���������, ���� �� ��������� ��������� ��������� ��������� Saved � ���� ��������� �� ���� ��������� - ��������� ��;
'If DocWord.Saved = False Then DocWord.Save


Set TableWord = DocWord.Tables(1)
'.Add(DocWord.Range(), 10, 2)


'�������� ����� � ������ � �������
'(�����_������, �����_�������)

TableWord.Cell(1, 2).Range.Text = MainForm.Label3
TableWord.Cell(2, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5)
TableWord.Cell(2, 2).Range.Text = "�� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)

TableWord.Cell(1, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)
TableWord.Cell(2, 3).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
TableWord.Cell(4, 1).Range.Text = "������ ���.��:" + MainForm.Label8 + "�."
TableWord.Cell(4, 2).Range.Text = Lic.Label10

'O9 = 0
'S9 = 0
'For rw = 1 To FG1.Rows - 1
'If FG1.TextMatrix(rw, 23) = "-" Then O9 = O9 + FG1.TextMatrix(rw, 18)
'If FG1.TextMatrix(rw, 23) = "s" Then S9 = S9 + FG1.TextMatrix(rw, 18)
'Next


TableWord.Cell(5, 2).Range.Text = Str(O9)
TableWord.Cell(6, 2).Range.Text = Str(S9)



TableWord.Cell(8, 1).Range.Text = "������ ���.��:" + MainForm.Label8 + "�."

TableWord.Cell(8, 2).Range.Text = Lic.Label13

End Sub

