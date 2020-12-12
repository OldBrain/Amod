VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexADO_DOMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' заказные источники данных FlexGrid осуществляют IVSFlexDataSource
Implements IVSFlexDataSource

' ОБЪЕКТЫ ХЛОПОТ и рекордного указателя
Public m_Conn As New ADODB.Connection
Public m_RS As New ADODB.Recordset
Dim m_Filter() As String
Dim m_lRecord As Long

'инициализируйте класс: создайте recordset
Private Sub Class_Initialize()
    m_Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
    m_Conn.Open "data/Kvartplata.mdb"
    m_RS.CursorLocation = adUseClient
    m_RS.Open "Kls_podr", m_Conn, adOpenDynamic
    ReDim m_Filter(m_RS.Fields.Count)
    m_lRecord = -1
End Sub

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String

    'первая запись является обычаем (запись поиска)
    If Record = 0 Then
        IVSFlexDataSource_GetData = m_Filter(Field)
        Exit Function
    End If
    
    ' курсор перемещения если мы имеем, чтобы
    If m_lRecord <> Record Then
        m_lRecord = Record
        m_RS.AbsolutePosition = Record
    End If
    
    ' величина retrieve
    If Not IsNull(m_RS.Fields(Field).Value) Then
        IVSFlexDataSource_GetData = m_RS.Fields(Field).Value
    End If
    
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = m_RS.Fields.Count
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    IVSFlexDataSource_GetFieldName = m_RS.Fields(Field).Name
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = m_RS.RecordCount + 1
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
    
    'мы только заинтересованы на линии фильтра
    If Record <> 0 Then Exit Sub
    
    ' сохраните новый пункт
    m_Filter(Field) = Trim(newData)
    If InStr("><=", Left(m_Filter(Field), 1)) = 0 Then
        m_Filter(Field) = "=" & m_Filter(Field)
    End If
        
    'постройте строку фильтра
    Dim i%, sFilter$
    For i = 0 To UBound(m_Filter)
        If Len(m_Filter(i)) Then
            If Len(sFilter) Then sFilter = sFilter & " AND "
            sFilter = sFilter & m_RS.Fields(i).Name & " "
            sFilter = sFilter & m_Filter(i)
        End If
    Next
        
    ' без изменения? никакая работа (отметьтесь, что преобразования ХЛОПОТ "" в 0)
    Dim sOldFilter$
    sOldFilter = m_RS.Filter
    If sOldFilter = "0" Then sOldFilter = ""
    If sFilter = sOldFilter Then Exit Sub
    
    ' строка фильтра правопреемника
    On Error Resume Next
    m_RS.Filter = sFilter
    If Err > 0 Then
        MsgBox "Error in filter string. Remember to enclose strings in single quotes (e.g. > 'BONAP').", vbInformation
        m_RS.Filter = ""
    End If
    m_RS.Requery
    On Error GoTo 0

End Sub
