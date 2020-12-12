VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form ERKC 
   Caption         =   "Импорт оплаты ЕРКЦ"
   ClientHeight    =   7464
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   10464
   LinkTopic       =   "Form3"
   ScaleHeight     =   7464
   ScaleWidth      =   10464
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить >>"
      Height          =   372
      Left            =   8280
      TabIndex        =   1
      Top             =   6960
      Width           =   1452
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5892
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10212
      _cx             =   18013
      _cy             =   10393
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "ERKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fname As String
Dim AccessRs As ADODB.Recordset
Dim rsDocReestr As ADODB.Recordset

Dim RS As ADODB.Recordset




Private Sub Command1_Click()
'Me.Enabled = False
ERKCVibor.Show 1
'MsgBox Str(ERKCVibor.code)

'MsgBox MainForm.Label8
Set rsDocReestr = New ADODB.Recordset

rsDocReestr.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.NachCod, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc"), Mconn, adOpenKeyset, adLockPessimistic
rsDocReestr.AddNew
rsDocReestr("Coment") = "Реестр ЕРКЦ" + Reestr
Cod = rsDocReestr("Cod")
rsDocReestr("Data") = Date
rsDocReestr("Nach") = "Любое начисление"
rsDocReestr.UpdateBatch
rsDocReestr.Close



' Добавляем в DOC

Mconn.Execute "INSERT INTO Doc ( Cod, KodKv, Summa, DataR, KodN, NameKv, Stst, Dom ) SELECT " + Str(Cod) + ", MainOccupant.Numer, Bank.SUMMA, Date(), '" + Str(ERKCVibor.code) + "', MainOccupant.FAM, 0 AS Выражение3, MainOccupant.Dom FROM Bank INNER JOIN MainOccupant ON Bank.NewNum = MainOccupant.BanKN"

Mconn.Execute "UPDATE doc INNER JOIN nachisleniy ON doc.KodN = nachisleniy.Kod SET doc.NameN = [nachisleniy]![Naim], doc.Tip = [nachisleniy]![Tip]"




'Выводим номера без соответствия

sq = "SELECT Bank.NewNum, Bank.SUMMA FROM Bank LEFT JOIN MainOccupant ON Bank.NewNum = MainOccupant.BanKN WHERE (((MainOccupant.BanKN) Is Null))"


Reports.sq = sq

Analizlgot.Titl = MainForm.Label3 + vbNewLine + " Эти номера лицевых счетов отсутствуют в базе!"

Analizlgot.Show
Analizlgot.fg1.DataRefresh
Unload Me


'**********************************



Unload ERKCVibor
Unload Me
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
    Set fg1.DataSource = RS

Set AccessRs = New ADODB.Recordset


Mconn.Execute "DELETE Bank.* FROM Bank"
AccessRs.Open ("SELECT Bank.DATA, Bank.KFOSB, Bank.SUMMA, Bank.SBOR, Bank.FIO, Bank.ADR, Bank.LSCHET, Bank.PERIODOPL, Bank.KV,Bank.LIFT, Bank.MUSOR, Bank.SELEN, Bank.GVoda, Bank.Otopl, Bank.HVoda, Bank.SSLIV, Bank.PLNOM, Bank.PLDATE, Bank.NRS, Bank.NewNum FROM Bank"), Mconn, adOpenDynamic, adLockBatchOptimistic


 



Do While Not RS.EOF

If Len(RS("Pay")) = 0 Then RS("Pay") = 0
'Pod.ProgressBar1.Index = I + 1

'If RS("Pay") <> 0 And Len(RS("Number")) = 12 Then

If RS("Pay") <> 0 Then


AccessRs.AddNew


AccessRs("Summa") = Replace(RS("Pay"), ".", ",") ' Общая сумма оплаты

'AccessRs("FIO") = RS("ФИО Ответственного") 'Фамилия имя отчество
'AccessRs("Adr") = RS("Адрес") 'Адрес

If RS("Number") = "" Then MsgBox RS("Number")

' Добавляем ведущие ноли до 12 символов
Num = Right(RS("Number"), 12)




Do While Len(Num) < 12
If Len(Num) < 12 Then Num = "0" + Num
Loop


AccessRs("NewNum") = Num ' Лиц.счет 12

'If RS("PERIODOPL") <> "" Then AccessRs("PERIODOPL") = RS("PERIODOPL") 'Период за который платит квартиросъемщик


If RS("Pay") <> "" Then AccessRs("KV") = Replace(RS("Pay"), ",", ".") 'КВАРТПЛАТА
'If RS("SLift") <> "" Then AccessRs("Lift") = Replace(RS("SLift"), ",", ".") 'ЛИФТ
'If RS("SMUSOR") <> "" Then AccessRs("MUSOR") = Replace(RS("SMUSOR"), ",", ".") 'МУСОР
'If RS("Selen") <> "" Then AccessRs("selen") = Replace(RS("selen"), ",", ".") '
'If RS("SGVS") <> "" Then AccessRs("Gvoda") = Replace(RS("SGVS"), ",", ".") 'ГОРЯЧАЯ ВОДА
'If RS("Steplo") <> "" Then AccessRs("Otopl") = Replace(RS("STEPLO"), ",", ".") 'ОТОПЛЕНИЕ
'If RS("SHVODA") <> "" Then AccessRs("HVoda") = Replace(RS("SHVODA"), ",", ".") 'ХОЛОДНАЯ ВОДА
'If RS("SSLIV") <> "" Then AccessRs("SSLIV") = Replace(RS("SSLIV"), ",", ".") 'СЛИВ
'AccessRs("PLNOM") = Replace(RS("PLNOM"), ",", ".") 'НОМЕР ПЛАТЕЖА
'AccessRs("PLDATE") = Replace(RS("PLDATE"), ",", ".") 'ДАТА ПЛАТЕЖА
'AccessRs("NRS") = Replace(RS("NRS"), ",", ".") ' Расчетный счет

AccessRs.UpdateBatch

End If

RS.MoveNext
Loop
   
    
    Set fg1.DataSource = AccessRs
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
