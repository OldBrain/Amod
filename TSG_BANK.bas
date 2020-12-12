Attribute VB_Name = "TSG_BANK"
Option Explicit
Dim RsTsg As ADODB.Recordset
Dim ConnTsg As ADODB.Connection

Public Sub TSGBank(myFld As ADODB.Field)
Set RsTsg = New ADODB.Recordset

For Each myFld In RsTsg.Fields
    'If myFld.Name = colname Then

MsgBox RsTsg.Fields


     '   Exit For
    'End If
Next



End Sub
