VERSION 5.00
Begin VB.Form Pocta 
   Caption         =   "Почта"
   ClientHeight    =   2424
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3216
   Icon            =   "Pocta.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2424
   ScaleWidth      =   3216
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Готово"
      Height          =   372
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "В реестре оплаты создан документ"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3012
   End
End
Attribute VB_Name = "Pocta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ShotFname As String
Public Fname As String
Public CodPL As String
Dim AccessRs As ADODB.Recordset
Dim rsDocReestr As ADODB.Recordset
Dim rsMain As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim NUM As String




Private Sub Command1_Click()
Unload Me


'**********************************





End Sub

Private Sub Form_Load()

'Pod.Show
DoEvents
'getXLData "C:/ERKC/ERKC.xls", "Лист1", "*", "", True
getXLData Fname, "Лист1", "*", "", True

'Unload Pod

End Sub


Sub getXLData(ByVal vstrWorkbookFullName As String, _
              ByVal vstrWorksheetName As String, _
              Optional ByVal vstrColumns As String = "*", _
              Optional ByVal vstrRange As String = "", _
              Optional ByVal vfUseHeader As Boolean)
    
    Const adOpenStatic = 3
    Const adLockReadOnly = 1
    
    Dim Conn As Object
    Dim RS As Object
    Dim strConnString As String
    Dim StrSQL As String
    'On Error GoTo HandleError
    Set Conn = CreateObject("ADODB.Connection")
    strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=" & IIf(vfUseHeader, "Yes", "No") & ";IMEX=1"";" _
                    & "Data Source=" & vstrWorkbookFullName
    Conn.Open strConnString
    
    Set RS = CreateObject("ADODB.Recordset")

    StrSQL = "SELECT " & vstrColumns & " FROM [" & vstrWorksheetName & "$" & vstrRange & "]"
    
        
    RS.Open StrSQL, Conn, adOpenStatic, adLockReadOnly
        
    'что-то читаем, передаем данные
'   Set FG1.DataSource = RS

Set AccessRs = New ADODB.Recordset
Set rsMain = New ADODB.Recordset


'Mconn.Execute "DELETE Bank.* FROM Bank"
'AccessRs.Open ("SELECT Bank.DATA, Bank.KFOSB, Bank.SUMMA, Bank.SBOR, Bank.FIO, Bank.ADR, Bank.LSCHET, Bank.PERIODOPL, Bank.KV,Bank.LIFT, Bank.MUSOR, Bank.SELEN, Bank.GVoda, Bank.Otopl, Bank.HVoda, Bank.SSLIV, Bank.PLNOM, Bank.PLDATE, Bank.NRS, Bank.NewNum FROM Bank"), Mconn, adOpenDynamic, adLockBatchOptimistic



'Добавляем строку в реестр документов

Set rsDocReestr = New ADODB.Recordset

rsDocReestr.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.NachCod, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc"), Mconn, adOpenKeyset, adLockPessimistic
rsDocReestr.AddNew
rsDocReestr("Coment") = "Реестр ПОЧТЫ " + Me.ShotFname
Cod = rsDocReestr("Cod")
rsDocReestr("Data") = Date
rsDocReestr("Nach") = "Любое начисление"
rsDocReestr.UpdateBatch
rsDocReestr.Close


'*********************************************
'Оскрываем строки документов для добавления
AccessRs.Open ("SELECT doc.Cod, doc.DataR, doc.KodN, doc.NameN, doc.KodKv, doc.NameKv, doc.Summa, doc.Key, doc.KeyAdding, doc.Stst, doc.Com, doc.Tip, doc.Button, doc.Dom, doc.RealData, doc.PLNOM FROM doc"), Mconn, adOpenDynamic, adLockBatchOptimistic
 



Do While Not RS.EOF


If Trim(RS("Лицевой счет")) <> "" Then

'MsgBox (RS("ФИО"))


If Len(RS("Сумма платежа")) = 0 Then RS("Pay") = 0
'Pod.ProgressBar1.Index = I + 1

'If RS("Pay") <> 0 And Len(RS("Number")) = 12 Then

If RS("Сумма платежа") <> 0 Then


AccessRs.AddNew





'Проверяем по какому счету связывать

Set rsSt = New ADODB.Recordset

rsSt.Open ("SELECT Settings.BankN, Settings.Neo FROM Settings"), Mconn

Neo = rsSt("Neo") ' код неопознанных сумм
If rsSt("BankN") <> "OLDNUM" Then

If RS("Лицевой счет") = "" Then MsgBox RS("Number")
' Добавляем ведущие ноли до 12 символов
NUM = Right(RS("Лицевой счет"), 12)
Do While Len(NUM) < 12
If Len(NUM) < 12 Then NUM = "0" + NUM
Loop


' Ишем абонента по 12 значному счету

rsMain.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.OLDNUM, MainOccupant.BanKN From MainOccupant WHERE (((MainOccupant.BanKN)='" + NUM + "'))")


Num1 = rsMain("Numer")


Else
NUM = Trim(RS("Лицевой счет"))




' Ишем абонента по OLDMUN  счету

rsMain.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.kv_num, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.OLDNUM, MainOccupant.BanKN From MainOccupant WHERE (((MainOccupant.OLDNUM)='" + NUM + "'))"), Mconn

End If


If rsMain.EOF = False Or rsMain.BOF = False Then 'Проверка что рекорсет не пустой

'Если лицевой счет найден************************************

Num1 = rsMain("Numer")

Else ' ************************ ЕСЛИ ЛИЦЕВОЙ СЧЕТ НЕ НАЙДЕН
Num1 = Neo


End If



' Добавляем строки документа

AccessRs("KodKv") = Num1 'Добавляем код абонента в реестр документов
AccessRs("Cod") = Cod
s = Val(RS("Сумма платежа"))
If RS("Сумма платежа") <> "" Then AccessRs("Summa") = s
'Replace(RS("Сумма платежа"), ",", ".") 'КВАРТПЛАТА
If Num1 <> Neo Then
AccessRs("NameKv") = rsMain("FAM") + " " + rsMain("Im") + " " + rsMain("Ot") 'Фамилия имя отчество
AccessRs("Dom") = rsMain("Dom") 'Код дома
Else
AccessRs("NameKv") = "НЕОПОЗНАННЫЕ СУММЫ"
AccessRs("Dom") = 0
End If
'AccessRs("Adr") = RS("Адрес") 'Адрес
AccessRs("PLNOM") = RS("Номер платежного поручения") 'Номер платежного поручения
AccessRs("Com") = "П/П № " + Str(RS("Номер платежного поручения")) + ". Оплата за " + CStr(RS("Период оплаты")) + " Дата оплаты: " + CStr(RS("Дата платежа")) + ". Оплатил:" + RS("ФИО") + " " + RS("Адрес") 'Коментарий
AccessRs("RealData") = RS("Дата платежа") 'Дата платежа
AccessRs("DataR") = RS("Дата платежа") 'Дата платежа
AccessRs("KodN") = Me.CodPL 'Код начисления


AccessRs.UpdateBatch





End If

'If Trim(RS("Лицевой счет")) <> "" Then

rsMain.Close

End If
RS.MoveNext



Loop
   
    
    
    AccessRs.Close
    '******************************
    
    
    RS.Close
    Set RS = Nothing
    
    
    
    Conn.Close
    Set Conn = Nothing
HandleExit:
    Exit Sub
HandleError:
    MsgBox "Error# " & Err.Number & vbCrLf & Err.Description
    Resume HandleExit
End Sub


Private Sub Проверка()

' Первоначальные цвета
fg1.Cell(flexcpForeColor, 1, 1, fg1.Rows - 1, fg1.Cols - 1) = vbBlack
fg1.Cell(flexcpFontBold, 1, 1, fg1.Rows - 1, fg1.Cols - 1) = False
fg1.Cell(flexcpBackColor, 1, 1, fg1.Rows - 1, fg1.Cols - 1) = vbWhite
fg1.Cell(flexcpBackColor, 1, 0, fg1.Rows - 1, 0) = &H8000000F


ErrorStst = 0


' Выделение ошибок цветом

For rw = 1 To fg1.Rows - 1

'MsgBox FG1.TextMatrix(Rw, 22)

For Cl = 1 To fg1.Cols - 1


'FG1.TextMatrix(Rw, Cl) = Replace(FG1.TextMatrix(Rw, Cl), ",", ".")

If fg1.TextMatrix(rw, Cl) = "" Then
fg1.TextMatrix(rw, Cl) = 0
fg1.Cell(flexcpForeColor, rw, Cl, rw, Cl) = vbRed
End If



Next Cl

'Ssum = Ssum + Round(Val(Replace(FG1.TextMatrix(rw, 4), ",", ".")), 2)


Next rw

'MsgBox Ssum

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Заполняем пустые строки DOC

Mconn.Execute ("UPDATE doc INNER JOIN nachisleniy ON doc.KodN = nachisleniy.Kod SET doc.NameN = [nachisleniy]![Naim], nachisleniy.Tip = [nachisleniy]![Tip]")

End Sub
