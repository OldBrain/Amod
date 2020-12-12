Attribute VB_Name = "Mod"
Option Explicit
Dim FunConn As ADODB.Connection
Dim FunRs As ADODB.Recordset



'Always on top

Declare Function SetWindowPos1 Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function Oplata(ByVal Num As String, Kat As String, Tip As String) As Double

Call BaseUnProtect(App.Path + "\data\" + "kvartplata.amd", True)

 Set FunConn = New ADODB.Connection
  FunConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.amd;Jet OLEDB:Database Password=" + MainForm.Pas + ";"
  FunConn.Open
  '"data/Kvartplata.mdb"
    
'Call BaseProtect(App.Path + "\data\" + "kvartplata.amd", True)
  
    
    
Set FunRs = New ADODB.Recordset
Set FunRs.ActiveConnection = FunConn
If Kat = "All" Then
FunRs.Open ("SELECT Adding.KodKv, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.Tip From Adding GROUP BY Adding.KodKv, Adding.Tip HAVING (((Adding.KodKv)=" + Num + ")  AND ((Adding.Tip)=" + Chr(34) + Tip + Chr(34) + "))")
Else
FunRs.Open ("SELECT Adding.KodKv, Adding.KodKat, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.Tip From Adding GROUP BY Adding.KodKv, Adding.KodKat, Adding.Tip HAVING (((Adding.KodKv)=" + Num + ") AND ((Adding.KodKat)=" + Kat + ") AND ((Adding.Tip)=" + Chr(34) + Tip + Chr(34) + "))")
End If
On Error GoTo Er
If Err.Number <> 3021 Then Oplata = FunRs("Sum-SummaI")
Er:
FunRs.Close
Set FunRs = Nothing
FunConn.Close
Set FunConn = Nothing
End Function

'Функция нечёткого сравнения использует в качестве аргументов две строки и
'параметр сравнения - максимальную длину сравниваемых подстрок.
'Результатом работы функции является число, лежащее в пределах от 0 до 1.
'0 соответствует полному несовпадению двух строк, а
'1 - полной (в определённом ниже смысле) их идентичности.
'
'Алгоритм
'Функция сравнения составляет все возможные комбинации подстрок
'с длинной вплоть до указанной (если длина 0 или есть строка с
'длинной меньше указанной длинны, то выбирается минимальная длина строк)
'и подсчитывает их совпадения в двух сравниваемых строках. Количество совпадений,
'разделённое на число вариантов объявляется коэффициентом схожести строк и
'выдаётся в качестве результата работы функции.



Public Function Compare(ByVal s As String, ByVal sOrig As String, Optional ByVal lLength As Long) As Double
    Compare = fStatL(s, sOrig, lLength)
End Function

Private Function fStatL(ByVal s As String, ByVal sOrig As String, ByVal lLength As Long) As Double
If Len(s) < lLength Or Len(sOrig) < lLength Or lLength = 0 Then If Len(s) < Len(sOrig) Then lLength = Len(s) Else lLength = Len(sOrig)

Dim L As Long
Dim lHits As Long
Dim lCount As Long
For L = 1 To lLength
    lHits = lHits + fStatIn(s, sOrig, L)
    lCount = lCount + Len(s) + Len(sOrig) - 2 * L + 2
Next L
If lCount = 0 Then fStatL = 0 Else fStatL = lHits / lCount
End Function

Private Function fStatIn(ByVal s As String, ByVal sOrig As String, ByVal lLength As Long) As Long
Dim L As Long
For L = 1 To Len(s) - lLength + 1
    If Not InStr(sOrig, Mid(s, L, lLength)) = 0 Then fStatIn = fStatIn + 1
Next L

For L = 1 To Len(sOrig) - lLength + 1
    If Not InStr(s, Mid(sOrig, L, lLength)) = 0 Then fStatIn = fStatIn + 1
Next L
End Function

