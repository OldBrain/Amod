VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form3"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim per As ADODB.Recordset
Dim nodess As Node
Set per = New ADODB.Recordset
КоннектЗ

per.Open ("SELECT Perecen.Код, Perecen.vid1, Perecen.vid2, Perecen.NameR, Perecen.sys FROM Perecen"), Zconn

'Set nodess = TreeView1.Nodes("a")



per.MoveFirst
Do While Not per.EOF



per.MoveNext
Loop
End Sub
