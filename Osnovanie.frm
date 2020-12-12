VERSION 5.00
Begin VB.Form Osnovanie 
   Caption         =   "Номер и основание платежа"
   ClientHeight    =   3876
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8976
   LinkTopic       =   "Form4"
   ScaleHeight     =   3876
   ScaleWidth      =   8976
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   4692
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4920
      TabIndex        =   4
      Top             =   3600
      Width           =   3972
   End
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   8892
   End
   Begin VB.ListBox List1 
      Height          =   1968
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   8772
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   0
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "П/ордер №"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1332
   End
End
Attribute VB_Name = "Osnovanie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 '#If Win32 Then  ' 32-разрядная версия VB
   Private Const LB_FINDSTRING = &H18F
   
   Private Declare Function SendMessage Lib _
     "user32" Alias "SendMessageA" (ByVal hwnd _
     As Long, ByVal wMsg As Long, ByVal wParam _
     As Long, lParam As _
     Any) As Long
    
    Dim It As String
Dim rsOsn  As ADODB.Recordset
    
 '#Else  ' 16-разрядная версия VB
  ' Private Const WM_USER = &H400
   'Private Const LB_FINDSTRING = (WM_USER + 16)
   'Private Declare Function SendMessage Lib _
    ' "User" (ByVal hWnd As Integer, ByVal wMsg _
     'As Integer, ByVal wParam As Integer, lParam _
     'As Any) As Long

Private Sub Command1_Click()
Me.Запись
Doc.Osnov = Me.Text1.Text
Doc.nu = Me.Text2.Text
Mconn.Execute ("UPDATE PO SET PO.Num = " + Text2.Text)
Unload Me
End Sub

Private Sub Command2_Click()
Doc.En = 10
Unload Me
End Sub

     '#End If
Private Sub Form_Load()


Set rsOsn = New ADODB.Recordset
rsOsn.Open ("SELECT PO.Num, PO.Основание FROM PO"), Mconn, adOpenStatic, adLockBatchOptimistic
Me.Text2.Text = rsOsn("Num") + 1


' Заполняем list
Me.Заполнитьсписок

'RsOsn.MoveFirst
'Do While Not RsOsn.EOF
'List1.AddItem (RsOsn("Основание"))
'RsOsn.MoveNext
'Loop


End Sub



Private Sub List1_DblClick()

Text1.Text = List1
End Sub

Private Sub Text1_Change()
  Dim pos As Long
  List1.ListIndex = SendMessage(List1.hwnd, _
    LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
  If List1.ListIndex = -1 Then
    pos = Text1.SelStart
  Else
    pos = Text1.SelStart
    Text1.Text = List1
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

Sub Заполнитьсписок()

'List1.AddItem "Апельсин"
 ' List1.AddItem "Банан"
 ' List1.AddItem "Яблоко"
 ' List1.AddItem "Персик"
 ' List1.AddItem "Ананас"
 ' List1.AddItem "Авокадо"
  

' Заполняем из рекордсета
rsOsn.MoveFirst
Do While Not rsOsn.EOF
List1.AddItem (rsOsn("Основание"))
rsOsn.MoveNext
Loop

End Sub

Sub Запись()
' Сначала проверяем есть ли уже такая запись
It = ""
rsOsn.MoveFirst
Do While Not rsOsn.EOF

If rsOsn("Основание") = Text1.Text Then
It = ""
' то выход из процедуры
'Loop
Exit Sub

Else
It = Text1.Text
End If

rsOsn.MoveNext
Loop

If It <> "" Then
rsOsn.AddNew
rsOsn("Основание") = It
rsOsn.UpdateBatch
rsOsn.Close


End If





End Sub

