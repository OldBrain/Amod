VERSION 5.00
Begin VB.Form BANKtxt 
   Caption         =   "������� � ����"
   ClientHeight    =   6900
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3528
   LinkTopic       =   "Form5"
   ScaleHeight     =   6900
   ScaleWidth      =   3528
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "��������(������ 9.1.) ��������� �����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   3132
   End
   Begin VB.CheckBox Check1 
      Caption         =   "12 ������� �.�����"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Value           =   1  'Checked
      Width           =   3132
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1800
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   4200
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������(������ 9.2) ���� ������."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   3132
   End
   Begin VB.DirListBox Dir1 
      Height          =   2232
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3132
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3132
   End
   Begin VB.Label Label3 
      Caption         =   "��������� ������� ��� ��������"
      Height          =   612
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1452
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000002&
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   3252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����������, ������� ���� ��� ������ �����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3492
   End
End
Attribute VB_Name = "BANKtxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName As String
Dim Period As String
Dim lineFile As String
Dim RsSet As ADODB.Recordset
Dim RsData As ADODB.Recordset
Dim rsCombo As ADODB.Recordset
Dim rsForSumm As ADODB.Recordset
Dim rsId As ADODB.Recordset ' ��������� �������� ��� ��� ����� ��� ��������
Dim Kategor As String




Private Sub Command1_Click()

' ������ ��������� 9.2 ��� ����� ������


If Me.Label2.Caption = "Label2" Then
MsgBox ("�� �� ������� ���� ��� ������ �����")
Exit Sub
End If




FileName = Label2.Caption
Label2.Caption = Replace(Label2.Caption, "\\", "\")
Kategor = Trim(Left(Combo1.Text, (InStr(Combo1.Text, " "))))


' ��������� ������ ��� lineFile
Set RsData = New ADODB.Recordset
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile(FileName, True)



RsData.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKat, Adding.Tip, KLS_PODR.[Imp] FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKat, Adding.Tip, KLS_PODR.[Imp] HAVING (((Adding.KodKat)=" + Kategor + ") AND ((Adding.Tip)='+') AND ((KLS_PODR.[Imp])=True))"), Mconn

'RsData.Open ("SELECT MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Sum(Adding.SummaI) AS [Sum-SummaI], Adding.KodKat, Adding.Tip, KLS_PODR.[Imp], Kategor.Nac FROM ((Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) INNER JOIN Kategor ON Adding.KodKat = Kategor.��� GROUP BY MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKat, Adding.Tip, KLS_PODR.[Imp], Kategor.Nac as KodUslugi HAVING (((Adding.KodKat)=" + Kategor + ") AND ((Adding.Tip)='+') AND ((KLS_PODR.[Imp])=True))"), Mconn

RsData.MoveFirst

Do While Not RsData.EOF
s = Replace(CStr(Format(RsData("Sum-SummaI"), "###0.00")), ",", ".")
'MsgBox (s)


If Me.Check1 Then ' ���� 12 ������� �����
lineFile = RsData("BanKN") + ";" + RsData("FAM") + " " + RsData("IM") + " " + RsData("OT") + ";" + "�.���������," + RsData("NAIM_KLS") + ",��.� " + RsData("kv_num") + ";" + Period + ";" + s
Else ' ���� 12 ������� �����
lineFile = RsData("OLDNUM") + ";" + RsData("FAM") + " " + RsData("IM") + " " + RsData("OT") + ";" + "�.���������," + RsData("NAIM_KLS") + ",��.� " + RsData("kv_num") + ";" + Period + ";" + s
End If
'Format(RsData("Sum-SummaI"), "###0.00")

a.WriteLine (lineFile)

RsData.MoveNext
Loop

a.Close

RsData.Close
RsSet.Close

Mconn.Execute ("UPDATE Settings SET Settings.Sh = 0 WHERE (((Settings.Sh)=999))")
Mconn.Execute ("UPDATE Settings SET Settings.Sh = [Settings]![Sh]+1")




'CreateAfile
MsgBox ("���� �������! >> " + FileName)
Unload Me
End Sub

Private Sub Command2_Click()
'������ ��������� 10 ���� ���������� �����
If Me.Label2.Caption = "Label2" Then

MsgBox ("�� �� ������� ���� ��� ������ �����")
Exit Sub
End If

FileName = Label2.Caption
Label2.Caption = Replace(Label2.Caption, "\\", "\")
Kategor = Trim(Left(Combo1.Text, (InStr(Combo1.Text, " "))))

' ��������� ������ ��� lineFile
Set RsData = New ADODB.Recordset
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile(FileName, True)






' ��������� �������� ��� ��� ����� ��� ��������



Set rsId = New ADODB.Recordset
rsId.Open ("SELECT MainOccupant.Numer From MainOccupant ORDER BY MainOccupant.Numer"), Mconn
', adOpenStatic, adLockBatchOptimistic

' ������� ������ ��� �����������
Fc = 0
rsId.MoveFirst
   Do While Not rsId.EOF
  Fc = Fc + 1
    
    rsId.MoveNext
Loop

'------------------------------

rsId.MoveFirst

t = 1






                    Do While Not rsId.EOF
    IdNum = Str(rsId("Numer"))
                    
                
                    DoEvents
                    
                    Pod.Show
                    Pod.ProgressBar1.Max = Fc + 10
                    Pod.ProgressBar1.Value = t
                    
                    
    ' ������� ������� ��� ������ ������ ������ �� rsId("Numer")
                    
                    
' RsData.Open ("SELECT MainOccupant.Numer, MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Max(Adding.SummaI) AS [Max-SummaI], Adding.KodKat, Adding.NameKat, Adding.Tip, KLS_PODR.[Imp] FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� GROUP BY MainOccupant.Numer, MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKat, Adding.NameKat, Adding.Tip, KLS_PODR.[Imp] HAVING (((MainOccupant.Numer)=" + IdNum + ") AND ((Adding.Tip)='+') AND ((KLS_PODR.[Imp])=True))"), Mconn
   
RsData.Open ("SELECT MainOccupant.Numer, MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Max(Adding.SummaI) AS [Max-SummaI], Adding.KodKat, Adding.NameKat, Adding.Tip, KLS_PODR.[Imp], Kategor.Nac FROM ((Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.���) INNER JOIN Kategor ON Adding.KodKat = Kategor.��� GROUP BY MainOccupant.Numer, MainOccupant.BanKN, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKat, Adding.NameKat, Adding.Tip, KLS_PODR.[Imp], Kategor.Nac HAVING (((MainOccupant.Numer)=" + IdNum + ") AND ((Adding.Tip)='+') AND ((KLS_PODR.[Imp])=True))"), Mconn

                                              '���� ��������� �� ������
    
                                        If RsData.EOF = False Or RsData.BOF = False Then
    RsData.MoveFirst
                                        
                    n = 0


Do While Not RsData.EOF

Pod.Label3.Caption = RsData("FAM") + " " + RsData("IM") + " " + RsData("OT") + " " + RsData("NAIM_KLS") + ",��.� " + RsData("kv_num")
Pod.Label1.Caption = "��� �� ������ ����� �������." + Chr(10) + " ��� ������������ " + Str(t) + "������� �� " + Str(Fc)
                                
n = n + 1

'��� ������ ����� �� ��������� Kategor ���� Nac

'MsgBox (Str(RsData("Nac")))



' ���������� ����� ���������� ��� ������ ��������� �������
Set rsForSumm = New ADODB.Recordset
NUM = Str(RsData("Numer"))
Kat = Str(RsData("KodKat"))
rsForSumm.Open ("SELECT Adding.KodKv, Adding.KodKat, Max(Adding.SummaI) AS [Max-SummaI] From Adding GROUP BY Adding.KodKv, Adding.KodKat HAVING (((Adding.KodKv)=" + NUM + ") AND ((Adding.KodKat)=" + Kat + "))"), Mconn
Sum = Replace(CStr(Format(rsForSumm("Max-SummaI"), "###0.00")), ",", ".")

rsForSumm.Close

s = Sum
'**************


If n = 1 Then ' ���� ��� ������ ��� �� ����� ��� � �.�.






's = Replace(CStr(Format(RsData("Max-SummaI"), "###0.00")), ",", ".")










If Me.Check1 Then ' ���� 12 ������� �����
lineFile = RsData("BanKN") + ";" + RsData("FAM") + " " + RsData("IM") + " " + RsData("OT") + ";" + "�.���������," + RsData("NAIM_KLS") + ",��.� " + RsData("kv_num") + ";" + Period + ";" + RsData("NameKat") + ";" + Str(RsData("Nac")) + ";" + s
Else ' ���� 12 ������� �����
lineFile = RsData("OLDNUM") + ";" + RsData("FAM") + " " + RsData("IM") + " " + RsData("OT") + ";" + "�.���������," + RsData("NAIM_KLS") + ",��.� " + RsData("kv_num") + ";" + Period + ";" + RsData("NameKat") + ";" + Str(RsData("Nac")) + ";" + s
End If
'Format(RsData("Sum-SummaI"), "###0.00")

Else ' ���� ����� ����� ��� ��� ������ ��������� ��� ������

lineFile = lineFile + ";" + RsData("NameKat") + ";" + Str(RsData("Nac")) + ";" + s

End If

                                    

RsData.MoveNext
                              Loop


'����� � ����
a.WriteLine (lineFile)



                                        End If '����� ���� ��������� �� ������
                                        
RsData.Close
                                        
                                        

   rsId.MoveNext
   
   t = t + 1

                    Loop
a.Close


RsSet.Close

Mconn.Execute ("UPDATE Settings SET Settings.Sh = 0 WHERE (((Settings.Sh)=999))")
Mconn.Execute ("UPDATE Settings SET Settings.Sh = [Settings]![Sh]+1")


Unload Pod

'CreateAfile
MsgBox ("���� �������! >> " + FileName)
Unload Me


End Sub

Private Sub Dir1_Change()
Dir1.Path = Drive1.Drive

Label2.Caption = Dir1.Path + FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveEr
Dir1.Path = Drive1.Drive

Label2.Caption = Dir1.Path + FileName
Label2.Caption = Replace(Label2.Caption, "\\", "\")
DriveEr:
If Err.Number = 68 Then MsgBox "��� ����� � ���������, ��� ���� ���������"
End Sub

Sub CreateAfile()
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set a = fs.CreateTextFile(FileName, True)
    
    a.WriteLine ("This is a test.")
    a.Close
End Sub

Private Sub Form_Load()
'Me.Label2.Caption = App.Path
'Me.Check1 = True
' ����� ������ ��� ������������ ����� �����
Set RsSet = New ADODB.Recordset


RsSet.Open ("SELECT Settings.TekData, Settings.Bank12, Settings.INN, Settings.Rs, Settings.Sh FROM Settings"), Mconn

Period = Mid(CStr(RsSet("TekData")), 4, 2) + Mid(CStr(RsSet("TekData")), 9, 2)



'D = Mid(CStr(Date), 4, 2)
D = Left(CStr(Date), 2)

sh = Right("00" + Trim(Str(RsSet("Sh"))), 3)


'***********************************************

'FileName = "\" + RsSet("INN") + "_" + RsSet("RS") + "_" + "" + sh + ".y" + D
FileName = "\" + RsSet("INN") + "_" + RsSet("RS") + "_" + "" + sh + ".txt"

'��������� ������� ��� ��������
Set rsCombo = New ADODB.Recordset
rsCombo.Open ("SELECT Kategor.���, Kategor.Name_Kategor FROM Kategor"), Mconn


rsCombo.MoveFirst
Do While Not rsCombo.EOF
Combo1.AddItem (CStr(rsCombo("���")) & "  " & rsCombo("Name_Kategor"))
'Combo1.AddItem (RsCombo("Name_Kategor"))
Combo1.ItemData(Combo1.NewIndex) = rsCombo("���")
rsCombo.MoveNext
Loop

rsCombo.MoveFirst
rsCombo.MoveNext
Combo1.Text = CStr(rsCombo("���")) & "  " & rsCombo("Name_Kategor")
rsCombo.Close
End Sub

