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
      Caption         =   "ОТМЕНА"
      Height          =   492
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   5652
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Заполнить справочную информацию для шапки квитанции"
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   5652
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Маленькая с информацией для сверки расчетов"
      Height          =   492
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   5652
   End
   Begin VB.CommandButton Command5 
      Caption         =   "За услуги лифта"
      Height          =   492
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   5652
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Начисления с общей задолженностью(портрет)"
      Height          =   492
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   5652
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Только начислено(портрет)"
      Height          =   492
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   5652
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Только начисления"
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   5652
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ФИО заполняет плательщик"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2892
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Печатать л/сч 12 знаков"
      Height          =   372
      Left            =   3480
      TabIndex        =   2
      Top             =   5280
      Value           =   1  'Checked
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "С оплатой и сальдо"
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
   'Описываем процедуры для вывода штрихкода
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
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'Dim WordApp1 As Word.Application ' экземпляр приложения
'Dim DocWord1 As Word.Document ' экземпляр документа
'Dim S As Integer
'Dim S1 As Integer

'объявляем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim i As Integer
Dim s As Double
'*****************************************


'Если не выбран адрес

If Combo1.Text = "Выбери адрес" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'Запоминаем код дома
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' Окно для выбора лицевых счетов

Me.Label1.Caption = fil

LSKvit.Show 1

'Если отмена то выходим
If Exit_Me = True Then Exit Sub



'MsgBox (fil)
'Описание Рекордсет для получения данных о начислениях
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'Получаем данные
'Блок получения номеров для одного дома
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")

rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")

'Получаем реквизиты для шапки
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'Описание Рекордсет для получения данных о начислениях по начислениям
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'Описание Рекордсет для получения данных о начислениях по КАТЕГОРИЯМ
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'Цикл по лицевым счетам дома
rsNum.MoveFirst
Do While Not rsNum.EOF




'Рекордсет для получения данных о начислениях одного лиц счета
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.TarifD, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** Запрос на выборку по категориям для заполнения строк квитанции
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.TarifD, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> 'ОДН')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "ПОЖАЛУЙСТА ПОДОЖДИТЕ. Сохраняю файлы квитанций."
Jdite.Label1 = rsNum("NAIM_KLS") + " Кв № " + RsKvit("kv_num") + " Лиц.счет" + rsNum("Oldnum")


' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'Задаем имя файла отчета
nameRP = "I"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + rsNum("NAIM_KLS") + "_Кв №_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'Убираем точку из названия файла
nameRP = Replace(nameRP, ".", "_")

'Убираем слэш из названия файла
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'активируем его
 DocWord.Activate

'Заполняем реквизиты
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

'Дата
TableWord.Cell(6, 1).Range.Text = "Расчетный период " + MainForm.Label8 + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If

' Адрес
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num")


' ФИО

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'Площадь
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'Прописано
TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'Оплата всего



OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
End If
OplataRS.Close




                            'проверяем что рекордсет не пустой
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'Цикл по начислениям одного лиц счета
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        'Объеденяем ячейки
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** Проставляем начисления
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "-"
    End If
    
    
    ' Объем услуг
    ' Если Parametr="прописано" то ставим прописано
    
    If RsKvitK("Parametr") = "прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' Если Parametr="прочие" то ставим прописано иначе *
    If RsKvitK("Parametr") = "прочие" Or RsKvitK("SchetZ") = "Пер" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' Если Parametr="счетчик" или "площадь" то ставим прописано иначе площадь
    If (RsKvitK("Parametr") = "площадь" Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    'Тариф
    
    If InStr(1, RsKvitK("NameN"), "найм") = 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    'Теаерь плата за найм
    
    If InStr(1, RsKvitK("NameN"), "найм") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   'Блок для общих начислений
        
    'S для подсчета сумы долка по строке
    s = 0
        
        If RsKvitK("SchetZ") = "Общие" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
                
     s = s + RsKvitK("SummaI")
        End If
        
    'Блок для ОДН начислений"
     If RsKvitK("SchetZ") = "ОДН" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     'Блок для ПЕРЕРАСЧЕТА начислений"
     If RsKvitK("SchetZ") = "Пер" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "-"
     
 ' End If
  
  
  'Нормативы потребления коммунальных услуг
  
     If RsKvitK("norm") <> 0 Then
     TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     Else
     TableWord.Cell(i, 8).Range.Text = "Х"
     End If
  
  
  'Показания приборов учета
     If RsKvitK("Sch") = "Да" Then
     If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
     If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " По нормативу"
     
     Else
     TableWord.Cell(i, 9).Range.Text = "-"
     End If
  
                                        
  'Расчеты по оплате на конец
  
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
        
        
                                   'ИТОГО
                                   
Set TableWord = DocWord.Tables(2)


'итого начислено
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 OplataRS.Close
 
 
 'Задолженность/переплата на начало периода
OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "Задолженность за прошлые периоды"
If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "Переплата за прошлые периоды "
If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

End If

  OplataRS.Close
 
 
 'ОПЛАЧЕНО
 
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
End If
OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 'Задолженность/переплата на конец периода она же и того к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
 End If
 OplataRS.Close
 
 
 
 ' итого к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 End If
 OplataRS.Close
        
        
        'проверяем что рекордсет не пустой
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "Формирование квитанций успешно завершено"


Unload Jdite

MsgBox ("Формирование квитанций успешно завершено. Файлы квитанций сохранены в " + App.Path + "\izv\")

Unload Reports
MainMenu.Enabled = True
Unload Me



End Sub

        
        
       

Private Sub Command2_Click()  ' Квитанция ТОЛЬКО ОПЛАТА

Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'Dim WordApp1 As Word.Application ' экземпляр приложения
'Dim DocWord1 As Word.Document ' экземпляр документа
'Dim s As Single
'Dim S1 As Integer

'объявляем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim nameRP As String
Dim s As Double
Dim i As Integer

'*****************************************


'Если не выбран адрес

If Combo1.Text = "Выбери адрес" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'Запоминаем код дома
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' Окно для выбора лицевых счетов

Me.Label1.Caption = fil

LSKvit.Show 1
'Если отмена то выходим
If Exit_Me = True Then Exit Sub







'MsgBox (fil)
'Описание Рекордсет для получения данных о начислениях
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'Получаем данные
'Блок получения номеров для одного дома
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")


'Получаем реквизиты для шапки
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'Описание Рекордсет для получения данных о начислениях по начислениям
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'Описание Рекордсет для получения данных о начислениях по КАТЕГОРИЯМ
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'Цикл по лицевым счетам дома
rsNum.MoveFirst
Do While Not rsNum.EOF




'Рекордсет для получения данных о начислениях одного лиц счета
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** Запрос на выборку по категориям для заполнения строк квитанции
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.TarifD, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> 'ОДН')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "ПОЖАЛУЙСТА ПОДОЖДИТЕ. Сохраняю файлы квитанций."
Jdite.Label1 = rsNum("NAIM_KLS") + " Кв № " + RsKvit("kv_num") + " Лиц.счет" + rsNum("Oldnum")


' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'Задаем имя файла отчета
nameRP = "ibn"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + rsNum("NAIM_KLS") + "_Кв №_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'Убираем точку из названия файла
nameRP = Replace(nameRP, ".", "_")

'Убираем слэш из названия файла
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'активируем его
 DocWord.Activate

'Заполняем реквизиты
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

'Дата
TableWord.Cell(6, 1).Range.Text = "Расчетный период " + MainForm.Label8 + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If

' Адрес
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num")


' ФИО

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'Площадь
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'Прописано
TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'Оплата всего В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА



'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close




                            'проверяем что рекордсет не пустой
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'Цикл по начислениям одного лиц счета
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        'Объеденяем ячейки
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** Проставляем начисления
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "X"
    End If
    
    
    ' Объем услуг
    ' Если Parametr="прописано" то ставим прописано
    
    If RsKvitK("Parametr") = "прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' Если Parametr="прочие" то ставим прописано иначе *
    If RsKvitK("Parametr") = "прочие" Or RsKvitK("SchetZ") = "Пер" Then
    TableWord.Cell(i, 3).Range.Text = "X"
    End If
    
    ' Если Parametr="счетчик" или "площадь" то ставим прописано иначе площадь
    If (RsKvitK("Parametr") = "площадь" Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    'Тариф
    
    If InStr(1, RsKvitK("NameN"), "найм") = 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    'Теаерь плата за найм
    
    If InStr(1, RsKvitK("NameN"), "найм") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   'Блок для общих начислений
        
    'S для подсчета сумы долка по строке
    s = 0
        
        If RsKvitK("SchetZ") = "Общие" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "X"
                
     s = s + RsKvitK("SummaI")
        End If
        
    'Блок для ОДН начислений"
     If RsKvitK("SchetZ") = "ОДН" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "X"
     s = s + RsKvitK("SummaI")
     End If
     
     'Блок для ПЕРЕРАСЧЕТА начислений"
     If RsKvitK("SchetZ") = "Пер" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "X"
     s = s + RsKvitK("SummaI")
     End If
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "X"
     
 ' End If
  
  
  'Нормативы потребления коммунальных услуг
  
     If RsKvitK("norm") <> 0 Then
     TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     Else
     TableWord.Cell(i, 8).Range.Text = "Х"
     End If
  
  
  'Показания приборов учета
     If RsKvitK("Sch") = "Да" Then
     If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
     If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " По нормативу"
     
     Else
     TableWord.Cell(i, 9).Range.Text = "Х"
     End If
  
                                        
  'Расчеты по оплате на конец
  
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
        
        
                                   'ИТОГО
                                   
Set TableWord = DocWord.Tables(2)


'итого начислено
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 OplataRS.Close
 
 
 'Задолженность/переплата на начало периода В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
 
'OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


'If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "Задолженность за прошлые периоды"
'If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "Переплата за прошлые периоды "
'If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



'TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

'End If

 ' OplataRS.Close
 
 
 'ОПЛАЧЕНО В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
 
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 'Задолженность/переплата на конец периода она же и того к оплате В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
'OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
' End If
 'OplataRS.Close
 
 
 
 ' итого к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 End If
 OplataRS.Close
        
        
        'проверяем что рекордсет не пустой
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "Формирование квитанций успешно завершено"


Unload Jdite

MsgBox ("Формирование квитанций успешно завершено. Файлы квитанций сохранены в " + App.Path + "\izv\")

Unload Reports
MainMenu.Enabled = True
Unload Me









End Sub

Private Sub Command3_Click() ' ПОРТРЕТ

Dim RsKvit As ADODB.Recordset
Dim rsNum As ADODB.Recordset
Dim RsRec As ADODB.Recordset
Dim RsKvitK As ADODB.Recordset
Dim OplataRS As ADODB.Recordset
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'Dim WordApp1 As Word.Application ' экземпляр приложения
'Dim DocWord1 As Word.Document ' экземпляр документа
'Dim S As Integer
'Dim S1 As Integer

'объявляем объектную переменную в разделе
' Generals формы
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


'Если не выбран адрес

If Combo1.Text = "Выбери адрес" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'Запоминаем код дома
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' Окно для выбора лицевых счетов

Me.Label1.Caption = fil

LSKvit.Show 1
'Если отмена то выходим
If Exit_Me = True Then Exit Sub







'MsgBox (fil)
'Описание Рекордсет для получения данных о начислениях
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'Получаем данные
'Блок получения номеров для одного дома
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")


'Получаем реквизиты для шапки
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'Описание Рекордсет для получения данных о начислениях по начислениям
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'Описание Рекордсет для получения данных о начислениях по КАТЕГОРИЯМ
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'Цикл по лицевым счетам дома
rsNum.MoveFirst
Do While Not rsNum.EOF




'Рекордсет для получения данных о начислениях одного лиц счета
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** Запрос на выборку по категориям для заполнения строк квитанции
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> 'ОДН')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "ПОЖАЛУЙСТА ПОДОЖДИТЕ. Сохраняю файлы квитанций."
Jdite.Label1 = rsNum("NAIM_KLS") + " Кв № " + RsKvit("kv_num") + " Лиц.счет" + rsNum("Oldnum")


' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'Задаем имя файла отчета
nameRP = "ipt"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + rsNum("NAIM_KLS") + "_Кв №_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'Убираем точку из названия файла
nameRP = Replace(nameRP, ".", "_")

'Убираем слэш из названия файла
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'активируем его
 DocWord.Activate

'Заполняем реквизиты
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(2, 2).Range.Text = MainForm.NamePr + ", ИНН:" + MainForm.INN + ", БАНК:" + MainForm.Bank + ", БИК:" + MainForm.BIK + ", кор.счет.:" + MainForm.KS + ", р.счет:" + MainForm.RS

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

'Дата
TableWord.Cell(5, 1).Range.Text = "Расчетный период " + MainForm.Label8 + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(1, 1).Range.Text = "Л.счет: " + RsKvit("BanKN")
Else
TableWord.Cell(1, 1).Range.Text = "Л.счет: " + RsKvit("oldnum")
End If

' Адрес

TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = "Адрес:" + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + ", площадь:" + Str(RsKvit("COMSPACE")) + "кв.м., прописано кол.чел.:" + Str(RsKvit("NLODGERF"))


'------*********------Формируем штрихкод-------------------------------
'Dim s As String
 
' Оаисание полей штрихкода
CodeVersion = "ST00011|" ' Сразу указываем CodePage =1 (WIN1251) и разделитель |
Name1 = "Name =" + Replace(MainForm.NamePr, Chr$(34), "'") + "|" 'Наименование получателя перевода Сразу заменяем кавычки на одинарные
PersonalAcc = "PersonalAcc=" + MainForm.RS + "|" 'Номер счета получателя перевода
BankName = "BankName =" + MainForm.Bank + "|" 'Наименование банка получателя перевода
BIC = "BIC = " + MainForm.BIK + "|" ' Ясно что бик
CorrespAcc = "CorrespAcc =" + MainForm.KS + "|" ' Корсчет
PayeeINN = "PayeeINN =" + MainForm.INN + "|" ' ИНН
Category = "Category =|" ' Необязательное поле можно оставить пустым
lastName = "lastName =" + RsKvit("FAM") + "|" 'Фамилия
firstName = "firstName =" + RsKvit("IM") + "|" 'Имя
middleName = "middleName =" + RsKvit("IM") + "|" ' Отчество

' Номер лицевого счета

If Me.Check1 Then
PersAcc = "PersAcc=" + RsKvit("BanKN") + "|"
Else
PersAcc = "PersAcc=" + RsKvit("oldnum") + "|"
End If

'Адрес КОНЕЦ ШТРИХКОДА СИМВОЛ РАЗДЕЛИТЕЛЯ НЕ ДОБАВЛЯЕМ
PayerAddress = "PayerAddress=" + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num")

' Формируем штрихкод полностью
s = CodeVersion + Name1 + PersonalAcc + BankName + BIC + CorrespAcc + PayeeINN + Category + lastName + firstName + middleName + PersAcc + PayerAddress




                      ' Не работает на Win10 только на XP поэтому меняем

' Обращаемся к объекту pdf417
'Set O = CreateObject("pdf417.clspdf417")
'b = O.pdf417(s, -1)
'Выводим в квитанцию
'TableWord.Cell(4, 2).Range.Text = b
'Set O = Nothing

'Новый код *****************

'txtPDF417.Text = ""
'txtPDF417.FontName = MW6PDF417R6.TTF
 'txtPDF417.FontName = cbxFontName.Text
 ' txtPDF417.FontSize = CInt(cbxFontSize.Text)
    
    ' encode string using PDF417
    
    
   ' Вызываем функцию формирования штрихкода
   ' S-Это строка которую кодируем
   ' 10,10 - Колонки столбцы влияет на размер штрихкода
    
    
    Call PDF417Encode(s, 2, _
                      2, 10, _
                      10, False, False)
    
    ' how many rows?
    RowCount = PDF417GetRows
    ' how many characters in one row?
    ColCount = PDF417GetCols
    
   
   ' PDF417GetCharAt возвращает штоих для шрифтов
    EncodedMsg = vbCrLf
    For i = 1 To RowCount
        For J = 1 To ColCount
            EncodedMsg = EncodedMsg & Chr(PDF417GetCharAt(i - 1, J - 1))
            'MsgBox (EncodedMsg)
        Next J
        EncodedMsg = EncodedMsg & vbCrLf
    Next i
   

'Выводим в квитанцию
TableWord.Cell(4, 2).Range.Text = EncodedMsg

'*************************



'------***********---------------------------------------

' ФИО

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")

End If

'Площадь
'TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'Прописано
'TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'Оплата всего В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА



'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close




                            'проверяем что рекордсет не пустой
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'Цикл по начислениям одного лиц счета
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        'Объеденяем ячейки
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** Проставляем начисления
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = " "
    End If
    
    
    ' Объем услуг
    ' Если Parametr="прописано" то ставим прописано
    
    If RsKvitK("Parametr") = "прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' Если Parametr="прочие" то ставим прописано иначе *
    If RsKvitK("Parametr") = "прочие" Or RsKvitK("SchetZ") = "Пер" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' Если Parametr="счетчик" или "площадь" то ставим прописано иначе площадь
    If (RsKvitK("Parametr") = "площадь" Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    'Тариф
    
    If InStr(1, RsKvitK("NameN"), "найм") = 0 Then
    
    ' Хотят пустые строки вместо нолей перепишнм поставим условие
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "-"
    End If
    
    'Теаерь плата за найм
    
    If InStr(1, RsKvitK("NameN"), "найм") <> 0 Then
    
    ' Хотят пустые строки вместо нолей перепишнм поставим условие
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    
    
    
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   'Блок для общих начислений
        
    'S для подсчета сумы долка по строке
    s = 0
        
        If RsKvitK("SchetZ") = "Общие" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
                
     s = s + RsKvitK("SummaI")
        End If
        
    'Блок для ОДН начислений"
     If RsKvitK("SchetZ") = "ОДН" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     'Блок для ПЕРЕРАСЧЕТА начислений"
     If RsKvitK("SchetZ") = "Пер" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     
     If s <> 0 Then TableWord.Cell(i, 7).Range.Text = Format(s, "0.00") Else TableWord.Cell(i, 7).Range.Text = " "
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = " "
     
 ' End If
  
  
  'Нормативы потребления коммунальных услуг В ЭТОЙ КВИТАНЦИИ НЕ НУЖНО
  
     'If RsKvitK("norm") <> 0 Then
     'TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     'Else
     'TableWord.Cell(i, 8).Range.Text = "Х"
     'End If
  
  
  'Показания приборов учета В ЭТОЙ КВИТАНЦИИ НЕ НУЖНО
   '  If RsKvitK("Sch") = "Да" Then
   '  If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
   '  If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " По нормативу"
     
   '  Else
   '  TableWord.Cell(i, 9).Range.Text = "Х"
   '  End If
  
                                        
  'Расчеты по оплате на конец В ЭТОЙ КВИТАНЦИИ НЕ НУЖНО
  
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
        
        
                                   'ИТОГО
                                   
Set TableWord = DocWord.Tables(2)


'итого начислено
 
' DocWord.Tables(2).Rows.Add
 'TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 'OplataRS.Close
 
 
 'Задолженность/переплата на начало периода В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
 
'OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


'If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "Задолженность за прошлые периоды"
'If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "Переплата за прошлые периоды "
'If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



'TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

'End If

 ' OplataRS.Close
 
 
 'ОПЛАЧЕНО В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
 
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 'Задолженность/переплата на конец периода она же и того к оплате В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
'OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
' End If
 'OplataRS.Close
 
 
 
 ' итого к оплате
OplataRS.Open ("SELECT Adding.KodKv, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.Tip From Adding GROUP BY Adding.KodKv, Adding.Tip HAVING (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.Tip)='+'))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 End If
 OplataRS.Close
        
        
        'проверяем что рекордсет не пустой
                 
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext
Loop


Jdite.Label1.Caption = "Формирование квитанций успешно завершено"


Unload Jdite

MsgBox ("Формирование квитанций успешно завершено. Файлы квитанций сохранены в " + App.Path + "\izv\")

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
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'Dim WordApp1 As Word.Application ' экземпляр приложения
'Dim DocWord1 As Word.Document ' экземпляр документа
'Dim S As Integer
'Dim S1 As Integer

'объявляем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim nameRP As String
Dim s As Double
Dim i As Integer

'*****************************************


'Если не выбран адрес

If Combo1.Text = "Выбери адрес" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If


'Запоминаем код дома
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' Окно для выбора лицевых счетов

Me.Label1.Caption = fil

LSKvit.Show 1
'Если отмена то выходим
If Exit_Me = True Then Exit Sub







'MsgBox (fil)
'Описание Рекордсет для получения данных о начислениях
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'Получаем данные
'Блок получения номеров для одного дома
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")


'Получаем реквизиты для шапки
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs FROM Settings")

'Описание Рекордсет для получения данных о начислениях по начислениям
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'Описание Рекордсет для получения данных о начислениях по КАТЕГОРИЯМ
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn



'Цикл по лицевым счетам дома
rsNum.MoveFirst
Do While Not rsNum.EOF




'Рекордсет для получения данных о начислениях одного лиц счета
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** Запрос на выборку по категориям для заполнения строк квитанции
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> 'ОДН')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "ПОЖАЛУЙСТА ПОДОЖДИТЕ. Сохраняю файлы квитанций."
Jdite.Label1 = rsNum("NAIM_KLS") + " Кв № " + RsKvit("kv_num") + " Лиц.счет" + rsNum("Oldnum")


' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'Задаем имя файла отчета
nameRP = "Ipt_z"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + rsNum("NAIM_KLS") + "_Кв №_" + RsKvit("kv_num") + "_" + rsNum("Oldnum")

'Убираем точку из названия файла
nameRP = Replace(nameRP, ".", "_")

'Убираем слэш из названия файла
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'активируем его
 DocWord.Activate

'Заполняем реквизиты
Set TableWord = DocWord.Tables(1)

'TableWord.Cell(1, 3).Select
TableWord.Cell(2, 2).Range.Text = MainForm.NamePr + ", ИНН:" + MainForm.INN + ", БАНК:" + MainForm.Bank + ", БИК:" + MainForm.BIK + ", кор.счет.:" + MainForm.KS + ", р.счет:" + MainForm.RS

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

'Дата
TableWord.Cell(5, 1).Range.Text = "Расчетный период " + MainForm.Label8 + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select

If Me.Check1 Then
TableWord.Cell(1, 1).Range.Text = "Л.счет: " + RsKvit("BanKN")
Else
TableWord.Cell(1, 1).Range.Text = "Л.счет: " + RsKvit("oldnum")
End If

' Адрес

TableWord.Cell(2, 1).Select
TableWord.Cell(2, 1).Range.Text = "Адрес:" + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + ", площадь:" + Str(RsKvit("COMSPACE")) + "кв.м., прописано кол.чел.:" + Str(RsKvit("NLODGERF"))

' ФИО

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'Площадь
'TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'Прописано
'TableWord.Cell(4, 6).Range.Text = RsKvit("NLODGERF")

'Оплата всего В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА



'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close




                            'проверяем что рекордсет не пустой
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'Цикл по начислениям одного лиц счета
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        'Объеденяем ячейки
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** Проставляем начисления
                           
                           
                           If RsKvitK("Tip") = "+" Then
  
  
   'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = " "
    End If
    
    
    ' Объем услуг
    ' Если Parametr="прописано" то ставим прописано
    
    If RsKvitK("Parametr") = "прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' Если Parametr="прочие" то ставим прописано иначе *
    If RsKvitK("Parametr") = "прочие" Or RsKvitK("SchetZ") = "Пер" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' Если Parametr="счетчик" или "площадь" то ставим прописано иначе площадь
    If (RsKvitK("Parametr") = "площадь" Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    'Тариф
    
    If InStr(1, RsKvitK("NameN"), "найм") = 0 Then
    
    ' Хотят пустые строки вместо нолей перепишнм поставим условие
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "-"
    End If
    
    'Теаерь плата за найм
    
    If InStr(1, RsKvitK("NameN"), "найм") <> 0 Then
    
    ' Хотят пустые строки вместо нолей перепишнм поставим условие
    If RsKvitK("Tarif") <> 0 Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("Tarif"), "0.00")
    Else
    TableWord.Cell(i, 4).Range.Text = " "
    End If
    
    
    
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    
      
   'Блок для общих начислений
        
    'S для подсчета сумы долка по строке
    s = 0
        
        If RsKvitK("SchetZ") = "Общие" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
                
     s = s + RsKvitK("SummaI")
        End If
        
    'Блок для ОДН начислений"
     If RsKvitK("SchetZ") = "ОДН" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     'Блок для ПЕРЕРАСЧЕТА начислений"
     If RsKvitK("SchetZ") = "Пер" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = " "
     s = s + RsKvitK("SummaI")
     End If
     
     
     If s <> 0 Then TableWord.Cell(i, 7).Range.Text = Format(s, "0.00") Else TableWord.Cell(i, 7).Range.Text = " "
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = " "
     
 ' End If
  
  
  'Нормативы потребления коммунальных услуг В ЭТОЙ КВИТАНЦИИ НЕ НУЖНО
  
     'If RsKvitK("norm") <> 0 Then
     'TableWord.Cell(i, 8).Range.Text = Str(RsKvitK("norm")) + "(" + RsKvitK("edizm") + ")"
     'Else
     'TableWord.Cell(i, 8).Range.Text = "Х"
     'End If
  
  
  'Показания приборов учета В ЭТОЙ КВИТАНЦИИ НЕ НУЖНО
   '  If RsKvitK("Sch") = "Да" Then
   '  If RsKvitK("nr") = False Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")"
         
   '  If RsKvitK("nr") Then TableWord.Cell(i, 9).Range.Text = Str(RsKvitK("Shc_new")) + "(" + RsKvitK("edizm") + ")" + " По нормативу"
     
   '  Else
   '  TableWord.Cell(i, 9).Range.Text = "Х"
   '  End If
  
                                        
  'Расчеты по оплате на конец В ЭТОЙ КВИТАНЦИИ НЕ НУЖНО
  
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
        
        
                                   'ИТОГО
                                   
Set TableWord = DocWord.Tables(2)


'итого начислено
 
' DocWord.Tables(2).Rows.Add
 'TableWord.Cell(1, 2).Range.Text = OplataRS("Sum-SummaI")
 End If
 'OplataRS.Close
 
 
 'Задолженность/переплата на начало периода В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
 
'OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

' If OplataRS.EOF = False Or OplataRS.BOF = False Then


'If (OplataRS("plus") + OplataRS("minus")) > 0 Then TableWord.Cell(2, 1).Range.Text = "Задолженность за прошлые периоды"
'If (OplataRS("plus") + OplataRS("minus")) < 0 Then TableWord.Cell(2, 1).Range.Text = "Переплата за прошлые периоды "
'If (OplataRS("plus") + OplataRS("minus")) = 0 Then TableWord.Cell(2, 1).Range.Text = "XXX"



'TableWord.Cell(2, 2).Range.Text = Format((OplataRS("plus") + OplataRS("minus")), "0.00")

'End If

 ' OplataRS.Close
 
 
 'ОПЛАЧЕНО В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
 
'OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
'If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
'Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
'End If
'OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 'Задолженность/переплата на конец периода она же и того к оплате В ЭТОЙ КВИТАНЦИИ НЕ НУЖНА
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


 TableWord.Cell(2, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")

'TableWord.Cell(2, 2).Range.Text = "БЕ БЕ БЕ"

 End If
OplataRS.Close
 
 
 
 ' итого к оплате
OplataRS.Open ("SELECT Adding.KodKv, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.Tip From Adding GROUP BY Adding.KodKv, Adding.Tip HAVING (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.Tip)='+'))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 TableWord.Cell(1, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 End If
 OplataRS.Close
        
        
        'проверяем что рекордсет не пустой
                 
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
       
       
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext
Loop


Jdite.Label1.Caption = "Формирование квитанций успешно завершено"


Unload Jdite

MsgBox ("Формирование квитанций успешно завершено. Файлы квитанций сохранены в " + App.Path + "\izv\")

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
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'Dim WordApp1 As Word.Application ' экземпляр приложения
'Dim DocWord1 As Word.Document ' экземпляр документа
'Dim S As Integer
'Dim S1 As Integer

'объявляем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim Spravka As String
Dim Pusto As String ' Для пустых квитанций
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


'Если не выбран адрес

If Combo1.Text = "Выбери адрес" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If

'Запоминаем код дома
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' Окно для выбора лицевых счетов

Me.Label1.Caption = fil

LSKvit.Show 1

'Если отмена то выходим
If Exit_Me = True Then Exit Sub

'MsgBox (fil)
'Описание Рекордсет для получения данных о начислениях
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'Получаем данные
'Блок получения номеров для одного дома
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")

rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")

'Получаем реквизиты для шапки
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs, Settings.Kvit FROM Settings")

'Описание Рекордсет для получения данных о начислениях по начислениям
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'Описание Рекордсет для получения данных о начислениях по КАТЕГОРИЯМ
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn




                                'проверяем что рекордсет не пустой
                            'If rsNum.EOF = False Or rsNum.BOF = False Then

'Цикл по лицевым счетам дома
rsNum.MoveFirst
Do While Not rsNum.EOF




'Рекордсет для получения данных о начислениях одного лиц счета
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** Запрос на выборку по категориям для заполнения строк квитанции
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.TarifD, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.LgotaVid, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> 'ОДН')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "ПОЖАЛУЙСТА ПОДОЖДИТЕ. Сохраняю файлы квитанций."
Jdite.Label1 = rsNum("NAIM_KLS") + " Кв № " + rsNum("kv_num") + " Лиц.счет" + rsNum("Oldnum")


' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'Задаем имя файла отчета
nameRP = "lift"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + rsNum("NAIM_KLS") + "_Кв №_" + rsNum("kv_num") + "_" + rsNum("Oldnum")

'Убираем точку из названия файла
nameRP = Replace(nameRP, ".", "_")

'Убираем слэш из названия файла
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

' Убираем * из названия файла
nameRP = Replace(nameRP, "*", "")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'активируем его
 DocWord.Activate

'Заполняем реквизиты
Set TableWord = DocWord.Tables(1)

'Заполняем справочную информацию
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

'Дата
TableWord.Cell(6, 1).Range.Text = "Расчетный период " + MainForm.Label8 + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select
' Для астраханьлифт номера всегда OLDNUM поэтому
Me.Check1 = False



                                   'проверяем что рекордсет не пустой
                            If RsKvit.EOF = False Or RsKvit.BOF = False Then

'If RsKvit("oldnum") = "" Then MsgBox ("")
If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If


' Адрес
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " Кв №" + rsNum("kv_num")





'------*********------Формируем штрихкод-------------------------------

 
 
 
 
' Оаисание полей штрихкода
CodeVersion = "ST00012|" ' Сразу указываем CodePage =1 (WIN1251) и разделитель |
Name1 = "Name=" + Replace(MainForm.NamePr, Chr$(34), "'") + "|" 'Наименование получателя перевода Сразу заменяем кавычки на одинарные
PersonalAcc = "PersonalAcc=" + MainForm.RS + "|" 'Номер счета получателя перевода
BankName = "BankName=" + MainForm.Bank + "|" 'Наименование банка получателя перевода
BIC = "BIC=" + MainForm.BIK + "|" ' Ясно что бик
CorrespAcc = "CorrespAcc=" + MainForm.KS + "|" ' Корсчет
PayeeINN = "PayeeINN=" + MainForm.INN + "|" ' ИНН
Category = "Category=|" ' Необязательное поле можно оставить пустым
lastName = "lastName=" + RsKvit("FAM") + "|" 'Фамилия
firstName = "firstName=" + RsKvit("IM") + "|" 'Имя
middleName = "middleName=" + RsKvit("OT") + "|" ' Отчество

' Номер лицевого счета

If Me.Check1 Then
PersAcc = "PersAcc=" + RsKvit("BanKN")
Else
PersAcc = "PersAcc=" + RsKvit("oldnum")
End If

'Адрес КОНЕЦ ШТРИХКОДА СИМВОЛ РАЗДЕЛИТЕЛЯ НЕ ДОБАВЛЯЕМ
PayerAddress = "PayerAddress=" + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + "|"

' Формируем штрихкод полностью
strQR = CodeVersion + Name1 + PersonalAcc + BankName + BIC + CorrespAcc
strQR = Trim(strQR)

strQR1 = PayeeINN + lastName + firstName + middleName + PayerAddress + PersAcc
strQR1 = Trim(strQR1)


' ФИО


If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'Площадь
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'Прописано
TableWord.Cell(4, 6).Range.Text = Str(RsKvit("NLODGER")) + "/" + Str(RsKvit("NLODGER"))

'Оплата всего


OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
SpravkaO = "Оплачено в текущем периоде:" + Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
Spravka = ""
End If
OplataRS.Close



                                 Else ' иначе выводим предупреждение о пустой квитанции
                                Pusto = Pusto + rsNum("NAIM_KLS") + " Кв №" + rsNum("kv_num") + Chr(13) + Chr(10)


                               End If





                            'проверяем что рекордсет не пустой
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'Цикл по начислениям одного лиц счета
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        'Объеденяем ячейки
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** Проставляем начисления
                           
                           
                           If RsKvitK("Tip") = "+" And RsKvitK("SummaI") <> 0 Then
  
  
   'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "-"
    End If
    
    
    ' Объем услуг
    ' Если LgotaVid="прописано" то ставим прописано
    
    If RsKvitK("LgotaVid") = "Прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' Если Parametr="прочие" то ставим прописано иначе *
    If RsKvitK("LgotaVid") = "Прочие" Or RsKvitK("SchetZ") = "Пер" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' Если Parametr="счетчик" или "площадь" то ставим прописано иначе площадь
    If (RsKvitK("LgotaVid") = "Общая пл." Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    'Тариф
    
    'Если ЛИФТ от прописанных
    If RsKvitK("LgotaVid") = "Прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("TarifD"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    'Если ЛИФТ от площади
    
    If RsKvitK("SchetZ") <> "Пер" Then
    
    If (RsKvitK("LgotaVid") = "Общая пл." Or RsKvitK("LgotaVid") = "Жил. пл.") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("Tarif")
    If (RsKvitK("LgotaVid") = "Прописано" Or RsKvitK("LgotaVid") = "Проживает") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
    
    End If
    
    'If (RsKvitK("LgotaVid") = "Общая пл." Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    'TableWord.Cell(I, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    'End If
    
    
      
   'Блок для общих начислений
        
    'S для подсчета сумы долга по строке
    s = 0
        
        If RsKvitK("SchetZ") = "Общие" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
            
     s = s + RsKvitK("SummaI")
        End If
        
    'Блок для ОДН начислений"
     If RsKvitK("SchetZ") = "ОДН" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     'Блок для ПЕРЕРАСЧЕТА начислений"
     If RsKvitK("SchetZ") = "Пер" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     'MsgBox (RsKvitK("SummaI") + "  " + Format(s, "0.00"))
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "-"
     
 
                                        
  'Расчеты по оплате на конец
  
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
        
        
        
        

        
                                   'ИТОГО
                                   
Set TableWord = DocWord.Tables(2)



'итого начислено
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
 
 'проверяем что рекордсет не пустой
'если пустой в конце выводим окно с предупреждением
  If OplataRS("Sum-SummaI") = 0 Then
  Pusto = Pusto + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + Chr(13) + Chr(10)
                            End If
 
 
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 3).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 
 ' Сумма для штрихкода
 
 
 SumQR = Format(OplataRS("Sum-SummaI"), "0.00")
 SumQR = Replace(SumQR, ".", "")
 SumQR = Replace(SumQR, ",", "")
 SumQR = "Sum=" + SumQR + "|"
 strQR = strQR + SumQR + strQR1
 
 'MsgBox (strQR)
 
 'OplataRS ("Sum-SummaI")
 SpravkaN = "Итого начислено в текущем периоде: " + Str(OplataRS("Sum-SummaI"))
 
 Else
 'проверяем что рекордсет не пустой
'если пустой в конце выводим окно с предупреждением
 
  Pusto = Pusto + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + Chr(13) + Chr(10)
                             
 End If
 OplataRS.Close
 
 
 
 'Задолженность/переплата на начало периода
OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


If (OplataRS("plus") + OplataRS("minus")) > 0 Then SpravkaZP = "Задолженность за прошлые периоды"
If (OplataRS("plus") + OplataRS("minus")) < 0 Then SpravkaZP = "Переплата за прошлые периоды "
If (OplataRS("plus") + OplataRS("minus")) = 0 Then SpravkaZP = ""



SpravkaZP = SpravkaZP + ": " + Format((OplataRS("plus") + OplataRS("minus")), "0.00")

End If

  OplataRS.Close
 
 
 'ОПЛАЧЕНО
 
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then

SpravkaO = "Поступила оплата в текущем периоде: " + Format(OplataRS("Sum-SummaI"), "0.00")

'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
SpravkaO = "Поступила оплата в текущем периоде: " + Format(0, "0.00")

End If
OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 'Задолженность/переплата на конец периода она же и того к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
 End If
 OplataRS.Close
 
 
 
 ' итого к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 'TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 
 If OplataRS("Sum-SK") < 0 Then SpravkaD = "Итого переплата на конец периода: " + Format(OplataRS("Sum-SK"), "0.00")
 If OplataRS("Sum-SK") >= 0 Then SpravkaD = "Итого долг на конец периода: " + Format(OplataRS("Sum-SK"), "0.00")
 
 End If
 OplataRS.Close
        
 
        
        
        'проверяем что рекордсет не пустой
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
'Set TableWord = DocWord.Tables(3)
       
      'Справочная информация
    'TableWord.Cell(2, 1).Range.Text = "Справочно(Для сверки расчетов):" + Chr(13) + Chr(10) + SpravkaZP + Chr(13) + Chr(10) + SpravkaO + Chr(13) + Chr(10) + SpravkaN + Chr(13) + Chr(10) + SpravkaD
    TableWord.Cell(2, 1).Range.Text = "Справочно(Для сверки расчетов): " + "" + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
    
    'TableWord.Cell(1, 1).Range.Text = "Справочно(Для сверки расчетов):" + "; " + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
       
       
       
'Вставляем картинку QR-Code
       
     
     
     
     
     
     
     
     
     
'"Привет Андрюха!" +
       
       
      strQR = Replace(strQR, " ", "")
      
    GenerateBMP StrPtr("C:\Example.bmp"), StrPtr(strQR), 3, 5, QualityLow
    
    
    DocWord.Shapes.AddPicture "C:\Example.bmp", , True, 235, 0, 100, 70
    
    
    
      
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "Формирование квитанций успешно завершено"


Unload Jdite

MsgBox ("Формирование квитанций успешно завершено. Файлы квитанций сохранены в " + App.Path + "\izv\")

If Len(Pusto) <> 0 Then

MsgBox ("ОБНАРУЖЕНЫ ПУСТЫЕ КВИТАНЦИИ" + Chr(13) + Chr(10) + Pusto)

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
' Блок описания переменных для вывода в World
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' экземпляр приложения
Dim DocWord As Word.Document ' экземпляр документа
'Dim WordApp1 As Word.Application ' экземпляр приложения
'Dim DocWord1 As Word.Document ' экземпляр документа
'Dim S As Integer
'Dim S1 As Integer

'объявляем объектную переменную в разделе
' Generals формы
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim Spravka As String
Dim Pusto As String ' Для пустых квитанций
Dim nameRP As String
Dim s As Double
Dim i As Integer
Dim SpravkaO As String
Dim SpravkaN As String
Dim SpravkaZP As String
Dim SpravkaD  As String
'*****************************************


'Если не выбран адрес

If Combo1.Text = "Выбери адрес" Then
Combo1.SetFocus
SendKeys "{F4}"
Exit Sub
End If

'Запоминаем код дома
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))

' Окно для выбора лицевых счетов

Me.Label1.Caption = fil

LSKvit.Show 1

'Если отмена то выходим
If Exit_Me = True Then Exit Sub

'MsgBox (fil)
'Описание Рекордсет для получения данных о начислениях
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn


'Получаем данные
'Блок получения номеров для одного дома
Set rsNum = New ADODB.Recordset
Set rsNum.ActiveConnection = Mconn
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")
'rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + "))")

rsNum.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.FLOOR, MainOccupant.OLDNUM, MainOccupant.Подразд, MainOccupant.Priv, MainOccupant.BanKN, MainOccupant.Dog, MainOccupant.podyezd, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.otm FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Dom)=" + Str(fil) + ") AND ((MainOccupant.otm)= True))")

'Получаем реквизиты для шапки
Set RsRec = New ADODB.Recordset
Set RsRec.ActiveConnection = Mconn
RsRec.Open ("SELECT Settings.Name, Settings.DolgnRuk, Settings.FIORuk, Settings.DolgnFin, Settings.FIOFin, Settings.DolgnOtv, Settings.FioOtv, Settings.Adres, Settings.Bank, Settings.BIK, Settings.INN, Settings.Ks, Settings.Rs, Settings.Kvit FROM Settings")

'Описание Рекордсет для получения данных о начислениях по начислениям
Set RsKvit = New ADODB.Recordset
Set RsKvit.ActiveConnection = Mconn

'Описание Рекордсет для получения данных о начислениях по КАТЕГОРИЯМ
Set RsKvitK = New ADODB.Recordset
Set RsKvitK.ActiveConnection = Mconn

Set OplataRS = New ADODB.Recordset
Set OplataRS.ActiveConnection = Mconn




                                'проверяем что рекордсет не пустой
                            'If rsNum.EOF = False Or rsNum.BOF = False Then

'Цикл по лицевым счетам дома
rsNum.MoveFirst
Do While Not rsNum.EOF




'Рекордсет для получения данных о начислениях одного лиц счета
RsKvit.Open ("SELECT Adding.KodKv, Adding.KodN, Adding.NameN, Adding.KodKat, Adding.NameKat, Adding.Tarif, Adding.Socmin, Adding.Propis, Adding.Projiv, Adding.ProLift, Adding.ObPl, Adding.PolPl, Adding.SummaI, Adding.SummaB, Adding.SaldoN, Adding.SaldoK, Adding.Tip, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, Adding.DnP, Adding.DnF, MainOccupant.* FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer Where (((Adding.KodKv) =" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat, Adding.Tip DESC")
'******** Запрос на выборку по категориям для заполнения строк квитанции
'RsKvitK.Open ("SELECT Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat WHERE (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") AND ((Adding.SummaI)<>0))")


'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0)) ORDER BY Adding.KodKat")

RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.TarifD, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.LgotaVid, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ")) ORDER BY Adding.KodKat")

'RsKvitK.Open ("SELECT Adding.Shc_new, Adding.Shc_old, Adding.NameN, Adding.TarifI, Adding.KodKv, Adding.NameKat, Adding.edizm, Kategor.Parametr, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.Tip, Saldo_Arh.SK, Adding.SaldoK, Adding.SchetZ, Adding.norm, Adding.Shc_new, Adding.Sch, Adding.KodKat, Adding.NameN, Adding.nr FROM Kategor INNER JOIN (Adding INNER JOIN Saldo_Arh ON (Adding.KodKat = Saldo_Arh.KodKat) AND (Adding.KodKv = Saldo_Arh.KodKV)) ON Kategor.Код = Saldo_Arh.KodKat Where (((Adding.KodKv)=" + Str(rsNum("Numer")) + ") And ((Adding.SummaI) <> 0) And (Adding.SchetZ<> 'ОДН')) ORDER BY Adding.KodKat")

'If RsKvitK.EOF = False Or RsKvitK.BOF = False Then

Jdite.Show

Jdite.Caption = "ПОЖАЛУЙСТА ПОДОЖДИТЕ. Сохраняю файлы квитанций."
Jdite.Label1 = rsNum("NAIM_KLS") + " Кв № " + rsNum("kv_num") + " Лиц.счет" + rsNum("Oldnum")


' Чистим табл. сальдо
Mconn.Execute ("DELETE Saldo.* FROM Saldo")
' Доб данные о конечном сальдо в табл. Saldo
Mconn.Execute ("INSERT INTO Saldo ( KodKV, KodKat, SK, SN ) SELECT Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.SaldoK, Adding.SaldoN")



'Задаем имя файла отчета
nameRP = "smol"
'создаём новый экземпляр Word-a
Set WordApp = New Word.Application

'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
WordApp.Visible = False


'*************************************
'// если нужно открыть имеющийся документ, то пишем такой код

Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'активируем его
DocWord.Activate
'сохраняем временный документ
nameRP = nameRP + rsNum("NAIM_KLS") + "_Кв №_" + rsNum("kv_num") + "_" + rsNum("Oldnum")

'Убираем точку из названия файла
nameRP = Replace(nameRP, ".", "_")

'Убираем слэш из названия файла
nameRP = Replace(nameRP, "/", "_")
nameRP = Replace(nameRP, "\", "_")

DocWord.SaveAs (App.Path + "\izv\" + nameRP)
DocWord.Close


'создаём новый экземпляр Word-a
'Set WordApp = New Word.Application
' Отключаем проверку орфографии для ускорения работы
WordApp.Options.CheckSpellingAsYouType = False

'// если нужно открыть имеющийся документ, то пишем такой код
Set DocWord = WordApp.Documents.Open(App.Path + "\izv\" + nameRP + ".doc")



'активируем его
 DocWord.Activate

'Заполняем реквизиты
Set TableWord = DocWord.Tables(1)



'Заполняем справочную информацию
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

'Дата
TableWord.Cell(6, 1).Range.Text = "Расчетный период " + MainForm.Label8 + " г."

'*************************************

'MsgBox (Str(rsNum("Numer")) + "  -    " + rsNum("Oldnum"))

'лицевой счет
'TableWord.Cell(2, 7).Select
' Для астраханьлифт номера всегда OLDNUM поэтому
'Me.Check1 = False



                                   'проверяем что рекордсет не пустой
                            If RsKvit.EOF = False Or RsKvit.BOF = False Then

'If RsKvit("oldnum") = "" Then MsgBox ("")
If Me.Check1 Then
TableWord.Cell(2, 7).Range.Text = RsKvit("BanKN")
Else
TableWord.Cell(2, 7).Range.Text = RsKvit("oldnum")
End If


' Адрес
'TableWord.Cell(3, 7).Select
TableWord.Cell(3, 7).Range.Text = rsNum("NAIM_KLS") + " Кв №" + rsNum("kv_num")


' ФИО

If Me.Check1 = False Then
'TableWord.Cell(4, 2).Select
TableWord.Cell(4, 2).Range.Text = RsKvit("FAM") + " " + RsKvit("IM") + " " + RsKvit("OT")
End If

'Площадь
TableWord.Cell(4, 4).Range.Text = RsKvit("COMSPACE")

'Прописано
TableWord.Cell(4, 6).Range.Text = Str(RsKvit("NLODGER")) + "/" + Str(RsKvit("NLODGER"))

'Оплата всего


OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then
'TableWord.Cell(4, 8).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
SpravkaO = "Оплачено в текущем периоде:" + Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(4, 8).Range.Text = Format(0, "0.00")
Spravka = ""
End If
OplataRS.Close



                                 Else ' иначе выводим предупреждение о пустой квитанции
                                Pusto = Pusto + rsNum("NAIM_KLS") + " Кв №" + rsNum("kv_num") + Chr(13) + Chr(10)


                               End If





                            'проверяем что рекордсет не пустой
                            If RsKvitK.EOF = False Or RsKvitK.BOF = False Then


i = 10

'Цикл по начислениям одного лиц счета
        RsKvitK.MoveFirst
        Do While Not RsKvitK.EOF
        
       'MsgBox (RsKvit("NameKat") + "  " + RsKvit("NameN") + " " + RsKvit("SchetZ"))
       
       
        'Объеденяем ячейки
        'DocWord.Tables(1).Rows(1).Cells(5).Select
        'DocWord.Tables(1).Range.Cells.Merge
        
      ' MsgBox (TableWord.Rows.Count)
                                    '****** Проставляем начисления
                           
                           
                           If RsKvitK("Tip") = "+" And RsKvitK("SummaI") <> 0 Then
  
  
   'Добавляем строку в таблицу
        DocWord.Tables(1).Rows.Add
                i = i + 1
        
    'TableWord.Cell(i, 11).Select
    'TableWord.Cell(i, 1).Range.Text = RsKvitK("NameKat")
    'MsgBox (TableWord.Rows.Count)
    
    TableWord.Cell(i, 1).Range.Text = RsKvitK("NameN")
    
    If RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 2).Range.Text = RsKvitK("edizm")
    Else
    TableWord.Cell(i, 2).Range.Text = "-"
    End If
    
    
    ' Объем услуг
    ' Если LgotaVid="прописано" то ставим прописано
    
    If RsKvitK("LgotaVid") = "Прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("Propis")
    End If
    
    ' Если Parametr="прочие" то ставим прописано иначе *
    If RsKvitK("LgotaVid") = "Прочие" Or RsKvitK("SchetZ") = "Пер" Then
    TableWord.Cell(i, 3).Range.Text = " "
    End If
    
    ' Если Parametr="счетчик" или "площадь" то ставим прописано иначе площадь
    If (RsKvitK("LgotaVid") = "Общая пл." Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 3).Range.Text = RsKvitK("ObPl")
    End If
    
    'Тариф
    
    'Если ЛИФТ от прописанных
    If RsKvitK("LgotaVid") = "Прописано" And RsKvitK("SchetZ") <> "Пер" Then
    TableWord.Cell(i, 4).Range.Text = Format(RsKvitK("TarifD"), "0.00")
    'If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    End If
    
    'Если ЛИФТ от площади
    
    If RsKvitK("SchetZ") <> "Пер" Then
    
    If (RsKvitK("LgotaVid") = "Общая пл." Or RsKvitK("LgotaVid") = "Жил. пл.") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("Tarif")
    If (RsKvitK("LgotaVid") = "Прописано" Or RsKvitK("LgotaVid") = "Проживает") Then TableWord.Cell(i, 4).Range.Text = RsKvitK("TarifI")
    
    End If
    
    'If (RsKvitK("LgotaVid") = "Общая пл." Or RsKvitK("Parametr") = "счетчик") And RsKvitK("SchetZ") <> "Пер" Then
    'TableWord.Cell(I, 4).Range.Text = RsKvitK("TarifI")
   ' If RsKvitK("TarifI") = 0 Then TableWord.Cell(i, 4).Range.Text = "X"
    'End If
    
    
      
   'Блок для общих начислений
        
    'S для подсчета сумы долга по строке
    s = 0
        
        If RsKvitK("SchetZ") = "Общие" Then
       'TableWord.Cell(i, 5).Range.Text = RsKvitK("SaldoN")
        If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
            
     s = s + RsKvitK("SummaI")
        End If
        
    'Блок для ОДН начислений"
     If RsKvitK("SchetZ") = "ОДН" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 5).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 5).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     'Блок для ПЕРЕРАСЧЕТА начислений"
     If RsKvitK("SchetZ") = "Пер" Then
     If RsKvitK("SummaI") <> 0 Then TableWord.Cell(i, 6).Range.Text = Format(RsKvitK("SummaI"), "0.00") Else TableWord.Cell(i, 6).Range.Text = "-"
     s = s + RsKvitK("SummaI")
     End If
     
     TableWord.Cell(i, 7).Range.Text = Format(s, "0.00")
     
     
     If TableWord.Cell(i, 7).Range.Text = "0.00" Then TableWord.Cell(i, 7).Range.Text = "-"
     
  
                                        
  'Расчеты по оплате на конец
  
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
        
        
        
        

        
                                   'ИТОГО
                                   
Set TableWord = DocWord.Tables(2)



'итого начислено
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 
 
 'проверяем что рекордсет не пустой
'если пустой в конце выводим окно с предупреждением
  If OplataRS("Sum-SummaI") = 0 Then
  Pusto = Pusto + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + Chr(13) + Chr(10)
                            End If
 
 
 
' DocWord.Tables(2).Rows.Add
 TableWord.Cell(1, 3).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
 'OplataRS ("Sum-SummaI")
 SpravkaN = "Итого начислено в текущем периоде: " + Str(OplataRS("Sum-SummaI"))
 
 Else
 'проверяем что рекордсет не пустой
'если пустой в конце выводим окно с предупреждением
 
  Pusto = Pusto + rsNum("NAIM_KLS") + " Кв №" + RsKvit("kv_num") + Chr(13) + Chr(10)
                             
 End If
 OplataRS.Close
 
 
 
 'Задолженность/переплата на начало периода
OplataRS.Open ("SELECT Saldo_Arh.KodKV, Sum(IIf([Saldo_Arh]![SK]>0,[Saldo_Arh]![SK],0)) AS plus, Sum(IIf([Saldo_Arh]![SK]<0,[Saldo_Arh]![SK],0)) AS minus From Saldo_Arh GROUP BY Saldo_Arh.KodKV HAVING (((Saldo_Arh.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


If (OplataRS("plus") + OplataRS("minus")) > 0 Then SpravkaZP = "Задолженность за прошлые периоды"
If (OplataRS("plus") + OplataRS("minus")) < 0 Then SpravkaZP = "Переплата за прошлые периоды "
If (OplataRS("plus") + OplataRS("minus")) = 0 Then SpravkaZP = ""



SpravkaZP = SpravkaZP + ": " + Format((OplataRS("plus") + OplataRS("minus")), "0.00")

End If

  OplataRS.Close
 
 
 'ОПЛАЧЕНО
 
OplataRS.Open ("SELECT Adding.Tip, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKv From Adding GROUP BY Adding.Tip, Adding.KodKv HAVING (((Adding.Tip)='-') AND ((Adding.KodKv)=" + Str(rsNum("Numer")) + "))")
If OplataRS.EOF = False Or OplataRS.BOF = False Then

SpravkaO = "Поступила оплата в текущем периоде: " + Format(OplataRS("Sum-SummaI"), "0.00")

'TableWord.Cell(3, 2).Range.Text = Format(OplataRS("Sum-SummaI"), "0.00")
Else
'TableWord.Cell(3, 2).Range.Text = Format(0, "0.00")
SpravkaO = "Поступила оплата в текущем периоде: " + Format(0, "0.00")

End If
OplataRS.Close
 
 
 'TableWord.Cell(3, 2).Range.Text =
 
 'Задолженность/переплата на конец периода она же и того к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(IIf([Saldo]![SK]>0,[Saldo]![SK],0)) AS plus, Sum(IIf([Saldo]![SK]<0,[Saldo]![SK],0)) AS minus From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then


' TableWord.Cell(3, 2).Range.Text = OplataRS("plus") + OplataRS("minus")
 'TableWord.Cell(2, 2).Range.Text = OplataRS("minus")
 End If
 OplataRS.Close
 
 
 
 ' итого к оплате
OplataRS.Open ("SELECT Saldo.KodKV, Sum(Saldo.SK) AS [Sum-SK] From Saldo GROUP BY Saldo.KodKV HAVING (((Saldo.KodKV)=" + Str(rsNum("Numer")) + "))")

 If OplataRS.EOF = False Or OplataRS.BOF = False Then
 

 'TableWord.Cell(4, 2).Range.Text = Format(OplataRS("Sum-SK"), "0.00")
 
 If OplataRS("Sum-SK") < 0 Then SpravkaD = "Итого переплата на конец периода: " + Format(OplataRS("Sum-SK"), "0.00")
 If OplataRS("Sum-SK") >= 0 Then SpravkaD = "Итого долг на конец периода: " + Format(OplataRS("Sum-SK"), "0.00")
 
 End If
 OplataRS.Close
        
 
        
        
        'проверяем что рекордсет не пустой
                 End If
               '  End If
            
        
        RsKvitK.Close
        RsKvit.Close
 
       
       
'Set TableWord = DocWord.Tables(3)
       
      'Справочная информация
    'TableWord.Cell(2, 1).Range.Text = "Справочно(Для сверки расчетов):" + Chr(13) + Chr(10) + SpravkaZP + Chr(13) + Chr(10) + SpravkaO + Chr(13) + Chr(10) + SpravkaN + Chr(13) + Chr(10) + SpravkaD
    TableWord.Cell(2, 1).Range.Text = "Справочно(Для сверки расчетов): " + "" + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
    
    'TableWord.Cell(1, 1).Range.Text = "Справочно(Для сверки расчетов):" + "; " + SpravkaZP + "; " + SpravkaO + "; " + SpravkaN + "; " + SpravkaD
       
       
'Сохраняем файл

DocWord.Save
 
DocWord.Close

WordApp.Quit

 


'определяем видимость Word-a по True - видимый,
'по False - не видимый (работает только ядро)
'WordApp.Visible = True


rsNum.MoveNext
Loop




Jdite.Label1.Caption = "Формирование квитанций успешно завершено"


Unload Jdite

MsgBox ("Формирование квитанций успешно завершено. Файлы квитанций сохранены в " + App.Path + "\izv\")

If Len(Pusto) <> 0 Then

MsgBox ("ОБНАРУЖЕНЫ ПУСТЫЕ КВИТАНЦИИ" + Chr(13) + Chr(10) + Pusto)

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




'Заполняем комбобокс
Dim Addrconn As ADODB.Recordset

Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic

Addrconn.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.Подразделение, KLS_PODR.Благ From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Combo1.Text = "Выбери адрес"


Addrconn.MoveFirst
Combo1.AddItem "Все дома"
Do While Not Addrconn.EOF
If Addrconn("КОД") <> -1 Then
Combo1.AddItem Trim(Str(Addrconn("КОД"))) + " " + Addrconn("NAIM_KLS") + " дом № " + Addrconn("Num")
End If
Addrconn.MoveNext
Loop
End Sub

Private Sub QR(s As String)
Dim O As Object
Dim a As String
'S = "Содержание и ремонт общего имущества МКД    м.кв.   67,3    8,30    558,59      558,59  Х   Х ХВС СЧЕТЧИК м.куб   10  17,10   171,00      171,00   8.03(м.куб)     767(м.куб) ОДН ХВС м.куб   67,3    0,00    19,17       19,17   Х   Х Слив СЧЕТЧИК    м.куб   10  17,85   178,50      178,50   8.03(м.куб)     767(м.куб) Электроэнергия счетчик  кВт 200 3,76    752,00      752,00   90(кВт)     14595(кВт) ОДН Электроэнергия  кВт.    67,3    0,00    59,74       59,74   Х   Х Домофон мес.    X   35,00   35,00       35,00   Х   Х Вывоз мусора    чел.    4   82,00   328,00      328,00  Х   Х"
Set O = CreateObject("pdf417.clspdf417")

Debug.Print O.pdf417(s, -1)
Set O = Nothing
End Sub

