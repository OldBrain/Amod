VERSION 5.00
Begin VB.Form Расчет2 
   Caption         =   "Расчет"
   ClientHeight    =   2484
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   2484
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Пожалуйста подождите. Идет расчет."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Расчет2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Isprav As Integer 'Isprav = 1 С учетом исправлений 2 Без учета испр
Dim R, n, KEY As Integer
'Dim mconn As ADODB.Connection
Dim Ras As ADODB.Recordset
Dim Formula As String



Private Sub Command1_Click()
Unload SposobR2
Unload Me
End Sub

Private Sub Form_Load()
Doc.Enabled = False
Command1.Enabled = False

'Set conn = New ADODB.Connection
'conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'  conn.Open "data/Kvartplata.mdb"
Set Ras = New ADODB.Recordset
Set Ras.ActiveConnection = Mconn
Ras.CursorType = adOpenForwardOnly
Ras.LockType = adLockBatchOptimistic
'Ras.Open ("Adding")
n = 0
i = 0
For R = 1 To Doc.Fg.Rows - 1

n = Val(Doc.Fg.TextMatrix(R, 5))
'MsgBox (Str(N))
Ras.Open ("SELECT Adding.KodKv, Adding.Formula, Adding.Key From Adding WHERE (((Adding.KodKv)=" + Str(n) + "))")

    On Error Resume Next
Ras.MoveFirst
Formula = "0"
KEY = 0
i = i + 1
'MsgBox (Str(r))
                                       Do While Not Ras.EOF

Formula = Ras.Fields("Formula").Value
KEY = Ras.Fields("Key").Value
'Isprav =  2 Без учета испр
If Isprav = 2 Then
'MsgBox (Str(KEY))
Mconn.Execute ("UPDATE Adding SET Adding.SummaI = " + Formula + ", Adding.Ispr = 0  WHERE (((Adding.key)=" + Str(KEY) + "))")
'Doc.FG.TextMatrix(FG.Row, 10) = 0

End If
'Isprav =  1 C учетом испр
If Isprav = 1 Then
Mconn.Execute ("UPDATE Adding SET Adding.SummaI = " + Formula + " WHERE (((Adding.key)=" + Str(KEY) + ") and (Adding.Ispr=0))")
End If
Ras.MoveNext
                                              Loop

Ras.Close
Next
Label1 = "Расчет окончен. Количество расчитанных лицевых счетов " + Str(i) + " из " + Doc.Label6


Mconn.Execute ("UPDATE Doc INNER JOIN Adding ON Doc.Key = Adding.KodDoc SET Doc.Summa = [Adding]![SummaI], Doc.Stst = [Adding]![ispr]")
Doc.Fg.Refresh


Command1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Doc.Enabled = True

End Sub




