VERSION 5.00
Begin VB.Form Z_New_Plan 
   Caption         =   "Ввод экон.обоснования"
   ClientHeight    =   7944
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11232
   LinkTopic       =   "Form4"
   ScaleHeight     =   7944
   ScaleWidth      =   11232
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   9240
      TabIndex        =   5
      Top             =   7680
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   7680
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6720
      Width           =   10812
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      Left            =   3480
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   120
      Width           =   6012
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   600
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1812
   End
   Begin VB.ListBox List1 
      Height          =   5808
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10932
   End
End
Attribute VB_Name = "Z_New_Plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************** взято с sql.ru
Option Explicit
 '#If Win32 Then  ' 32-разрядная версия VB
   Private Const LB_FINDSTRING = &H18F
   
   Private Declare Function SendMessage Lib _
     "user32" Alias "SendMessageA" (ByVal hwnd _
     As Long, ByVal wMsg As Long, ByVal wParam _
     As Long, lParam As _
     Any) As Long
'***********************************************************
Dim RsPer As ADODB.Recordset ' Для полного перечня услуг подставляем в List
Dim Comb1RsPer As ADODB.Recordset 'Для заполнения Combo1
Dim Comb2RsPer As ADODB.Recordset 'Для заполнения Combo2
Dim It As String



Private Sub Combo1_Click()

' После выбора комбо заполняем комбо 2
' сначала очищаем
Me.Combo2.Clear
' заполняем комбо 2
Comb2RsPer.Open ("SELECT Perecen.vid2 From Perecen GROUP BY Perecen.vid1, Perecen.vid2 HAVING (((Perecen.vid1)='" + Combo1.Text + "'))"), Mconn
Comb2RsPer.MoveFirst
Me.Combo2.Text = "*"
Do While Not Comb2RsPer.EOF
Me.Combo2.AddItem (Comb2RsPer("vid2"))
Comb2RsPer.MoveNext
Loop
Comb2RsPer.Close

End Sub



Private Sub Combo2_Click()
' После выбора комбо2 заполняем listbox
Me.Заполнить

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Me.Запись

' Добавляем новую строчку
ZPlan.Fg.AddItem "+"
' FG  строчка, колонка
ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 1) = ZPlan.Dom

'ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 3) = ZPlan.T

ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 3) = ZPlan.T
ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 4) = Me.Combo1.Text

ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 5) = Me.Combo2.Text
ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 6) = Me.Text1.Text
ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 7) = 0
ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 8) = 0

ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 10) = Trim(ZPlan.Combo1.Text)
ZPlan.Fg.TextMatrix(ZPlan.Fg.Rows - 1, 11) = ZPlan.Combo2.Text

'ZPlan.Fg.Refresh

Unload Me

End Sub

Private Sub List1_DblClick()
Text1.Text = List1
Me.Text1.FontBold = True
End Sub

Private Sub Form_Load()
Set RsPer = New ADODB.Recordset
Set Comb1RsPer = New ADODB.Recordset
Set Comb2RsPer = New ADODB.Recordset

Me.Combo2.Text = "*"


Comb1RsPer.Open ("SELECT Perecen.vid1 From Perecen GROUP BY Perecen.vid1"), Mconn
'Заполняем комбо1
Comb1RsPer.MoveFirst
Me.Combo1.Text = "*"
Do While Not Comb1RsPer.EOF
Me.Combo1.AddItem (Comb1RsPer("vid1"))
Comb1RsPer.MoveNext
Loop
Comb1RsPer.Close



'RsPer.Open (""), Mconn
End Sub
Sub Заполнить()
' сначала очищаем List
Me.List1.Clear
' Открываем  рекордсет
RsPer.Open ("SELECT Perecen.kod, Perecen.vid1, Perecen.vid2, Perecen.NameR, Perecen.sys From Perecen WHERE (((Perecen.vid2)='" + Combo2.Text + "'))"), Mconn

' Заполняем listbox из рекордсета
RsPer.MoveFirst
Do While Not RsPer.EOF
List1.AddItem (RsPer("NameR"))
RsPer.MoveNext
Loop
' Закрываем рекордсет
RsPer.Close
End Sub

Sub Запись()
'MsgBox (RsPer.State)
 If RsPer.State = 0 Then
' Открываем  рекордсет для записи если он еще не открыт т.е. RsPer.State = 0
RsPer.Open ("SELECT Perecen.kod, Perecen.vid1, Perecen.vid2, Perecen.NameR, Perecen.sys From Perecen WHERE (((Perecen.vid2)='" + Combo2.Text + "'))"), Mconn, adOpenStatic, adLockBatchOptimistic
End If


' Сначала проверяем есть ли уже такая запись
It = ""
RsPer.MoveFirst
Do While Not RsPer.EOF

If RsPer("NameR") = Text1.Text Then
It = ""
' то выход из процедуры
'Loop
Exit Sub

Else
It = Text1.Text
End If

RsPer.MoveNext
Loop


' Собственно запись
                            If It <> "" Then
                        
                        
If MsgBox("Добавить новую запись <<" + It + ">> в справочник затрат?", vbYesNo) = vbYes Then

                        
RsPer.AddNew
RsPer("NameR") = It
RsPer("Vid1") = Me.Combo1.Text
RsPer("Vid2") = Me.Combo2.Text
RsPer("Sys") = 0
RsPer.UpdateBatch
RsPer.Close

End If ' End if msgbox
                                End If





End Sub



'******************** взято с sql.ru
Private Sub Text1_Change()
  Dim pos As Long
  List1.ListIndex = SendMessage(List1.hwnd, _
    LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
  If List1.ListIndex = -1 Then
    pos = Text1.SelStart
    Me.Text1.FontBold = False
  Else
    pos = Text1.SelStart
    Text1.Text = List1
    
    Me.Text1.FontBold = True
    
    Text1.SelStart = pos
    Text1.SelLength = Len(Text1.Text) - pos
    End If
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '
  On Error Resume Next
  If KeyCode = 8 Then ' Backspace
    If Text1.SelLength <> 0 Then
      Text1.Text = Mid$(Text1, 1, Text1.SelStart - 1)
      KeyCode = 0
    End If
  ElseIf KeyCode = 46 Then ' Del
     If Text1.SelLength <> 0 And _
       Text1.SelStart <> 0 Then
       KeyCode = 0
    End If
  End If
End Sub



