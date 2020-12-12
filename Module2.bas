Attribute VB_Name = "Module2"
'Вот еще вариант,
'(1) в отдельный модуль поместить:

'Public Sub FindInCombo(cbo As ComboBox)
'(C)Pinchuk Vitaly & Gura Nataly
'#Created: 05.02.04 - 13:13
'#Modifyed: 19.02.04 - 09:55

'Dim i As Integer, InputLen As Integer
'If cbo.Text = "" Then
'Exit Sub
'End If
'For i = 0 To cbo.ListCount
'If InStr(UCase(cbo.List(i)), UCase(cbo.Text)) = 1 Then
'InputLen = Len(cbo.Text)
'cbo.Text = cbo.List(i)
'cbo.SelStart = InputLen
'cbo.SelLength = Len(cbo.List(i)) - InputLen
'End If
'Next i
'End Sub

