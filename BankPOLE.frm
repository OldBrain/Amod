VERSION 5.00
Begin VB.Form BankPOLE 
   Caption         =   "Form4"
   ClientHeight    =   4992
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9972
   LinkTopic       =   "Form4"
   ScaleHeight     =   4992
   ScaleWidth      =   9972
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Связать поля"
      Height          =   612
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   2412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Вабрать файл банка для настройки"
      Height          =   612
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   2412
   End
End
Attribute VB_Name = "BankPOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBFConn As ADODB.Connection
'Dim mconn As ADODB.Connection
Dim dbfRs As ADODB.Recordset
Dim EtalonRs As ADODB.Recordset
Public DBFName As String

Private Sub Command1_Click()
BankImport.Command2.Visible = False
BankImport.Command3.Visible = True
BankImport.Show 1
End Sub

Private Sub Command2_Click()
Set DBFConn = New ADODB.Connection
'DBFName = BankImport.File1.FileName

MsgBox (Me.DBFName)
If Me.DBFName <> "" Then
DBFConn.Open "Provider=MSDASQL.1;Persist Security Info=False;mode=19;Data Source=Файлы dBASE;Initial Catalog=" + BankImport.File1.Path
End If

End Sub

