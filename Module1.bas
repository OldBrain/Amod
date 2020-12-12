Attribute VB_Name = "Module1"
Option Explicit

Type AllKvResult
    nKey As Long
    Familia As String
    name As String
    SecondName As String
End Type

Dim RsSet As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim mCol As Collection
Dim cOsn As ADODB.Connection
'Dim Conn As ADODB.Connection
Public Sub ConnectToBase(ByVal ConnString As String)
   Set cOsn = New ADODB.Connection
  cOsn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Jet OLEDB:Database Password=;"
  cOsn.Open
  GetDan
   End Sub


Public Sub ConnectArhive(ByVal ConnString As String)
   Set mCol = New Collection
  'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Jet OLEDB:Database Password=" + MainForm.Pas + ";"
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "янв"))
   'MsgBox Str(Year(RsSet("TekData"))) + "янв"
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Фев"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Мар"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Апр"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Май"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Июн"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Июл"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Авг"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Сен"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Окт"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Ноя"))
   Add ConnString, Trim(UCase(Str(Year(RsSet("TekData"))) + "Дек"))
   '...
End Sub


Private Sub Add(ByVal ConnString As String, ByVal CustomBase As String)
   Dim ConnX As ADODB.Connection
   Set ConnX = New ADODB.Connection
   ConnX.CursorLocation = adUseClient
   On Error GoTo NoFile

NoFile:
'MsgBox Err.Description
If Err.Number <> -2147467259 Then

   ConnX.Open UCase(Replace(ConnString, "/kvartplata.mdb", "/arhiv/" + Trim(UCase(CustomBase)) + ".mdb"))
   
   'MsgBox "'" + Replace(ConnString, "/kvartplata.mdb", "/arhiv/" + Trim(UCase(CustomBase)) + ".mdb") + "'"
   
   Err.Clear
   End If
   'MsgBox ConnX
   mCol.Add ConnX, CustomBase 'вот здесь
  
End Sub


    
 Public Property Get Conn(ByVal ConnName As String) As ADODB.Connection
'  ConnName = Trim(UCase(ConnName))
  Set Conn = mCol(ConnName)
End Property
 
Public Sub GetDan()
Set RsSet = New ADODB.Recordset
RsSet.Open ("SELECT Settings.* FROM Settings"), cOsn
End Sub


'Павел

'Public Function AllKv(nKey As Long, indata As AllKvResult) As AllKvResult


   ' null
   ' empy
    ' nothing
    
    
    'Dim result As AllKvResult
    'Dim Rs1 As ADODB.recodset
    'Set rs1 = New ADODB.Recordset
    'RsSet.Open ("SELECT MainOccupant.*, MainOccupant.Numer From MainOccupant WHERE (((MainOccupant.Numer)='+nKey+'))"), cOsn
    
    'result.Familia = Rs_Add.Fields("FAM").Value
    'result.name = Rs1.felds(3).Value
  
    'AllKv = result
'End Function



