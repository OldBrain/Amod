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
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
Dim TableWord As Word.Table
'*****************************************


'Получаем данные

'Задаем имя файла отчета
nameRP = "I"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + "1.doc"

DocWord.SaveAs (App.Path + "\" + nameRP)
DocWord.Close

'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\" + nameRP)

'активируем его
 DocWord.Activate
'Заполняем реквизиты
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

'Дата
TableWord.Cell(6, 1).Range.Text = "Расчетный период " + "MainForm.Label8" + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select
TableWord.Cell(2, 7).Range.Text = "RsKvit(BanKN)"
' Адрес
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = "rsNum(NAIM_KLS)"

' ФИО
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = "RsKvit(FAM)"

'Площадь
TableWord.Cell(4, 4).Range.Text = "RsKvit(COMSPACE)"

'Прописано
TableWord.Cell(4, 6).Range.Text = "RsKvit(NLODGERF)"

'Оплата всего





TableWord.Cell(4, 8).Range.Text = Format("100", "0.00")

TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")






                            'проверяем что рекордсет не пустой
                            



'Цикл по начислениям одного лиц счета
       
       For i = 11 To 23
        
       
       
                    
                                    '****** Проставляем начисления
     'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
      TableWord.Cell(i, 1).Range.Text = "NameN"
      TableWord.Cell(i, 2).Range.Text = "edizm"
    
    ' Объем услуг
     TableWord.Cell(i, 3).Range.Text = 10
    'Тариф
    TableWord.Cell(i, 4).Range.Text = Format(10, "0.00")
    s = 0
      
   TableWord.Cell(i, 5).Range.Text = Format(10, "0.00")
   TableWord.Cell(i, 6).Range.Text = Format(10, "0.00")
   TableWord.Cell(i, 7).Range.Text = Format(10, "0.00")
  
  'Нормативы потребления коммунальных услуг
     TableWord.Cell(i, 8).Range.Text = "Х"
     
  'Показания приборов учета
     TableWord.Cell(i, 9).Range.Text = "Х"
  
  'Расчеты по оплате на конец
    s = 100
  
     TableWord.Cell(i, 10).Range.Text = Format(s, "0.00")
     
     If TableWord.Cell(i, 10).Range.Text = TableWord.Cell(i - 1, 10).Range.Text Then
              
     TableWord.Cell(i, 10).Merge MergeTo:=TableWord.Cell(i - 1, 10)
   End If
     
  
       
                                         Next i
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext





Unload Jdite


End Sub

