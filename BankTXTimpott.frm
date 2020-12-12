VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BankTXTimpott 
   AutoRedraw      =   -1  'True
   Caption         =   "Импорт текстовых файлов *h "
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
      Caption         =   "Открыть файл"
      Height          =   492
      Left            =   2280
      TabIndex        =   6
      Top             =   5040
      Width           =   5532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отмена"
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
      Caption         =   "Укажите путь к файлу банка. Текстовый файл по маске  *.??y"
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


Command2.Caption = "Ждите идет обработка файла"


Dim rsReestrDoc As ADODB.Recordset



fileN = Trim(Me.Label2.Caption)

Set FSO = CreateObject("Scripting.FileSystemObject")
Set F = FSO.OpenTextFile(fileN)
 
 

 
'Считываем файл построчно
Do While Not F.AtEndOfStream
   Stroka = F.ReadLine
   TestArray = Split(Stroka, ";")
     
 ' перебераем строки для поиска последней строки с "=" и формируем заголовок реестра
     
    For I = 0 To UBound(TestArray)
    If TestArray(I) <> "" Then
    '    LastNonEmpty = LastNonEmpty + 1
'        TestArray(LastNonEmpty) = TestArray(i)
        
If InStr(TestArray(I), "=") = 1 Then ' Признак последней строки файла

Kol = Replace(TestArray(I), "=", "") ' количество строк в реестре

'Назначаем прогресбар
Me.ProgressBar1.Max = Kol + 10
Me.ProgressBar1.Value = 1
Me.ProgressBar1.Visible = True


Sum = TestArray(1) 'Общая сумма принятых средств
npp = TestArray(4) 'Номер платежного поручения
DPP = Left(TestArray(5), 10) 'Дата платежного поручения

'Формируем строку описания реестра
DataPP = CDate(DPP)
KomR = "Принято " + Sum + " Руб. № п/п " + npp + " от " + DPP + " кол.платежей=" + Kol

'MsgBox (KomR)
End If
      
      
      'MsgBox (TestArray(i))
    End If
Next
Loop

KodN = Trim(Left(Combo1.Text, (InStr(Trim(Combo1.Text), " "))))

'Добавляем строку в реестр документов
Mconn.Execute ("INSERT INTO ReestrDoc ( Data, NachCod, Nach, Coment, Summa, Status, Tip, KodDom, Adres ) SELECT '" + Replace(DPP, "-", "/") + "', " + KodN + ", 'Банк','" + KomR + "' , 0, 0, 'Реестр банка', 0, 'Все адреса'")


    'Mconn.Execute ("INSERT INTO ReestrDoc ( Data, Coment ) SELECT '" + Replace(DPP, "-", "/") + "', '" + KomR + "'")
F.Close
    
    'Находим код нового документа
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
    
    
    
                          'Добавляем строки платежей
   'Считываем файл построчно
 Set F = FSO.OpenTextFile(fileN)
                    Do While Not F.AtEndOfStream
                    
RaznesenoOK = False 'статус того что счет найден в базе
                    

                    
   Stroka = F.ReadLine
   TestArray = Split(Stroka, ";")
     
     
     
 ' перебераем строки
     
    'For i = 0 To UBound(TestArray)
    If TestArray(0) <> "" And InStr(TestArray(0), "=") = 0 Then
    
    Dpl = Trim(TestArray(0)) ' Дата платежа
    Filial = TestArray(2) 'Номер отделения
    Ls = Trim(TestArray(5)) ' Лицевой счет
        'Проверяем счет на кол-во символов и ведущие ноли
        '
            
            
     Me.Label4.Caption = "Разношу л/сч. " + Trim(TestArray(5)) + " ФИО " + TestArray(6) + " Адрес:" + TestArray(7)
     
     
            
            If Me.Label3.Caption = "BanKN" Then  ' если связь по 12-значному номеру
            Do While Len(Trim(Ls)) < 12
    Ls = "0" + Ls
            Loop
           
         Else
' Так же если по OLDNUM то убераем ведущий ноль
            
Do While Left(Ls, 1) = 0 ' убираем Все ведущие ноли
            
  If Left(Ls, 1) = 0 Then
  Ls = Right(Ls, Len(Ls) - 1)
  'MsgBox ("Внимание! Убрали ведущие ноли из л.сч. " + Ls + ".")
  End If
            Loop
            
           
            
            End If
            
Dim rsMain As ADODB.Recordset
'Ищем абонента в базе
Set rsMain = New ADODB.Recordset


' Слишком долго ремим код до звездочек и пробуем ускорить

rsMain.Open ("SELECT MainOccupant.Numer, KLS_PODR.КОД AS Dom, MainOccupant.OLDNUM, MainOccupant.BanKN FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД"), Mconn

'RSmain.Open ("SELECT MainOccupant.Numer, KLS_PODR.КОД AS Dom, MainOccupant.OLDNUM, MainOccupant.BanKN FROM MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД WHERE (((MainOccupant.Numer)=" + Ls + "))"), Mconn








rsMain.MoveFirst
Do While Not rsMain.EOF
DoEvents

'****************************************************
    If Me.Label3.Caption = "BanKN" Then  ' если связь по 12-значному номеру
  'То присваеваем LSN MainOccupant.BanKN
    Lsn = rsMain("BanKN")
    Else
  'То присваеваем LSN MainOccupant.OLDNUM
    Lsn = rsMain("OLDNUM")
  'Заодно убераем ведущие ноли
  Do While Left(Lsn, 1) = 0 ' убираем Все ведущие ноли
            
  If Left(Lsn, 1) = 0 Then
  Lsn = Right(Lsn, Len(Lsn) - 1)
  'MsgBox ("Внимание! Убрали ведущие ноли из л.сч. " + Ls + ".")
  End If
            Loop
  
  
  
    End If

'****************************************************






  If Lsn = Ls Then
  
  
  
  
  'If rsMain(Trim(Me.Label3.Caption)) = Ls Then
  LS1 = Trim(Str(rsMain("Numer")))
  Dom = Trim(Str(rsMain("Dom")))
  Ls = ""
  RaznesenoOK = True ' Ставим статус разнесен
  End If
   rsMain.MoveNext
   Loop
    
' ****************************************************************
    
    FIO = TestArray(6) 'ФИО
    Add = TestArray(7) 'Адрес
    NOtd = TestArray(2) 'Номер отделения
    nkass = TestArray(3) 'Номер кассира/УС/СБОЛ
    Period = Trim(TestArray(8)) ' Период оплаты
    Summ = TestArray(9) 'Сумма операции
    Summ = Replace(Summ, ",", ".")
    SummKomis = TestArray(11) ' Сумма комиссии банка
    
    
    
    
    
 Komm = "за " + Left(Period, 2) + "-" + Right(Period, 2) + ". Пл/п №" + Trim(npp) + " Отд.банка№" + Trim(NOtd) + "Кассир№" + Trim(nkass)
 
 ' Пофиксил вставку периода оплпты, банк может дать любые символы в дате преобрахование невозможно
 'If Len(Period) <> 0 Then
 'Period = "01/" + Left(Period, 2) + "/" + Right(Period, 2) '  период оплаты
 'Else
 'Period = "01/01/01"
 'End If
 Period = "01/01/01"
 
 
 If RaznesenoOK Then

 
 Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS Выражение1, " + KodN + " AS Выражение2, " + LS1 + " AS Выражение3, '" + FIO + "' AS Выражение4, " + Summ + " AS Выражение5, '-' AS Выражение6, " + npp + " AS Выражение7, '" + Replace(Dpl, "-", "/") + "' AS Выражение8, '" + Komm + "' AS Выражение9, " + Dom + " AS Выражение10, '" + Period + "' AS Выражение11")
    
 Else
  
  
If Dom = "" Then Dom = "-1"
  
Komm = "л/сч. " + Trim(TestArray(5)) + " ФИО " + TestArray(6) + " Адрес:" + TestArray(7) + "Реестр от " + Dpl
Mconn.Execute ("INSERT INTO Doc ( Cod, KodN, KodKv, NameKv, Summa, Tip, PLNOM, DataR, Com, Dom, RealData ) SELECT " + Trim(Cod) + " AS Выражение1, " + KodN + " AS Выражение2, " + Me.Label5.Caption + " AS Выражение3, '" + "Неопознанные суммы" + "' AS Выражение4, " + Summ + " AS Выражение5, '-' AS Выражение6, " + npp + " AS Выражение7, '" + Replace(Dpl, "-", "/") + "' AS Выражение8, '" + Komm + "' AS Выражение9, " + Dom + " AS Выражение10, '" + Period + "' AS Выражение11")
If Not RaznesenoOK And Ls <> "" Then MsgBox ("Внимание! Абанент с лиц счетом " + Ls + "не найден в базе. Фамлмия " + FIO + " Адрес " + Add + " Сумма " + TestArray(9) + ". РАЗНОШУ В НЕОПОЗНАННЫЕ СУММЫ.")
  
  
 End If
  
  
    End If
    
    
    
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    
    

    
    
   
                    Loop
       
F.Close

' Заполняем пустые стороки реестра

Mconn.Execute ("UPDATE doc INNER JOIN nachisleniy ON doc.KodN = nachisleniy.Kod SET doc.NameN = [nachisleniy]![Naim] WHERE (((doc.NameN) Is Null))")

MsgBox ("Данные разнесены. В реестре создан документ №" + Cod)
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
If Err.Number = 68 Then MsgBox "Нет диска в дисководе, или диск поврежден"
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
Dim fileN As String ' Путь и имя файла для чтения
Dim Stroka As String ' Считываемая строка
Dim DataPP As Date
Dim rsCombo As ADODB.Recordset
Dim RsSet As ADODB.Recordset
Dim RaznesenoOK As Boolean

Dim Ls As String 'Лицевой счет банка
Dim TipLs As String 'Поле по которому привязывается лиц. счет из табл. Setting
Dim NEOPOZNAN As String 'Номер л/сч неопознанных сумм


'Определяем какой номер лиц. счета для связи с банком
Set RsSet = New ADODB.Recordset
RsSet.Open ("SELECT Settings.BankN, Settings.Neo FROM Settings"), Mconn
TipLs = Trim(RsSet("BankN"))
NEOPOZNAN = Trim(RsSet("Neo"))
Me.Label3.Caption = TipLs
RsSet.Close
Me.Label5.Caption = NEOPOZNAN
LastNonEmpty = -1
File1.Pattern = "*.y??;*.txt"
'Заполняем комбобокс
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
