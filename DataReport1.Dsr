VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} DataReport123 
   Caption         =   "DataReport1"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   14526
   _Version        =   393216
   _DesignerVersion=   100685828
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   1
   GridY           =   1
   LeftMargin      =   1440
   RightMargin     =   1440
   TopMargin       =   1440
   BottomMargin    =   1440
   NumSections     =   3
   SectionCode0    =   2
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section2"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode1    =   4
   BeginProperty Section1 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section1"
      Object.Height          =   1440
      NumControls     =   1
      ItemType0       =   3
      BeginProperty Item0 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "Label1"
         Object.Width           =   2268
         Object.Height          =   240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Label1"
      EndProperty
   EndProperty
   SectionCode2    =   7
   BeginProperty Section2 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section3"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
End
Attribute VB_Name = "DataReport123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_kat As ADODB.Recordset
Dim TheConn As ADODB.Connection

Private Sub DataReport_Initialize()
' open connection
   Set TheConn = New ADODB.Connection
  TheConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  TheConn.Open "data/Kvartplata.mdb"


Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = TheConn
 
Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic
Rs_kat.Open "KLS_PODR"


Set DataReport1.DataSource = Rs_kat
Set Rs_kat = New ADODB.Recordset
Set Rs_kat.ActiveConnection = TheConn
 
Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic
Rs_kat.Open "KLS_PODR"
End Sub
