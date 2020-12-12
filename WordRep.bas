Attribute VB_Name = "WordRep"
Option Explicit

Public Sub ќтчет(nameRP As String)

Dim WordApp As Word.Application ' экземпл€р приложени€
Dim DocWord As Word.Document ' экземпл€р документа
'объ€вл€ем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long

'nameRP = "lc.doc"

'создаЄм новый экземпл€р Word-a
Set WordApp = New Word.Application

'определ€ем видимость Word-a по True - видимый,
'по False - не видимый (работает только €дро)
WordApp.Visible = True

'создаЄм новый документ в Word-e
'Set DocWord = WordApp.Documents.Add

'// если нужно открыть имеющийс€ документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP)

'активируем его
DocWord.Activate

'сохран€ем временный документ
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)
'ѕроверить, были ли сохранены внесенные изменени€ свойством Saved и если изменени€ не были сохранены - сохранить их;
'If DocWord.Saved = False Then DocWord.Save


Set TableWord = DocWord.Tables(1)
'.Add(DocWord.Range(), 10, 2)


'печатаем текст в €чейке с адресом
'(номер_строки, номер_столбца)

TableWord.Cell(1, 2).Range.Text = MainForm.Label3
TableWord.Cell(2, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5)
TableWord.Cell(2, 2).Range.Text = " в є" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)

TableWord.Cell(1, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)
TableWord.Cell(2, 3).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
TableWord.Cell(4, 1).Range.Text = "—альдо нач.на:" + MainForm.Label8 + "г."
TableWord.Cell(4, 2).Range.Text = Lic.Label10

'O9 = 0
'S9 = 0
'For rw = 1 To FG1.Rows - 1
'If FG1.TextMatrix(rw, 23) = "-" Then O9 = O9 + FG1.TextMatrix(rw, 18)
'If FG1.TextMatrix(rw, 23) = "s" Then S9 = S9 + FG1.TextMatrix(rw, 18)
'Next


TableWord.Cell(5, 2).Range.Text = Str(O9)
TableWord.Cell(6, 2).Range.Text = Str(S9)



TableWord.Cell(8, 1).Range.Text = "—альдо кон.на:" + MainForm.Label8 + "г."

TableWord.Cell(8, 2).Range.Text = Lic.Label13

End Sub

