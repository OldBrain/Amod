Attribute VB_Name = "OstrovODRepMdl"
Option Explicit

Type odrdate
    'What number, which evaluated as first, is with indifference
    wyear As Integer
    wmonth As Integer
    wday As Integer
End Type


'struct inhabitantsstruct {
'    odrdate datain;
'    char * fio;
'    short birthyear;
'    char * relationship;
'};

Type inhabitantsstruct
    datain As odrdate
    FIO As String
    birthyear As Integer
    relationship As String
End Type


'struct ostrovodrinstruct {
'    double area;
'    char * entitycardnum;
'    char * entityfio;
'    char * street;
'    char * house;
'    char * flat;
'    char * order;
'    char * orgname
'    char * regionname
'};

Public Type ostrovodrinstruct
    area As Double
    entitycardnum As String
    entityfio As String
    street As String
    house As String
    flat As String
    order As String
    orgname As String
    regionname As String
End Type

Public odrin As ostrovodrinstruct
Public inhs() As inhabitantsstruct


'int WINAPI DLL_EXPORT ostrovodrep(const ostrovodrinstruct * oodrin, const char* cszworkdir, HWND _hOwner, const char* cszZipCmd = "7z a -tzip")

Declare Function ostrovodrep Lib "ostrovodrep.dll" _
    Alias "ostrovodrep@16" (oodrin As ostrovodrinstruct, ByVal cszworkdir As String, ByVal hOwner As Long, Optional ByVal cszZipCmd As String) As Long

Declare Sub addinhabitant Lib "ostrovodrep.dll" _
    Alias "addinhabitant@4" (ByRef inhabitant As inhabitantsstruct)

Declare Sub deleteinhabitants Lib "ostrovodrep.dll" _
    Alias "deleteinhabitants@0" ()


Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long


Public rsODRGlobal As ADODB.Recordset

Public Sub ostrovodrepinit(frmowner As Form, RSInputData As ADODB.Recordset)
    
    Set rsODRGlobal = RSInputData
    
    Load frmOstrovodrepInhs
    
    Set frmOstrovodrepInhs.frmowner = frmowner
    'Set frmOstrovodrepInhs.RSInputData = RSInputData
    
    frmOstrovodrepInhs.Show 1, frmowner
End Sub












