VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������"
   ClientHeight    =   5340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "������� ���������� ���� � ����� KV.EXE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Pt As String


Private Sub CancelButton_Click()
Dialog.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
Mass.Show
Mass.Refresh
Mass.Label2.Caption = "���������, ���� ��������� ������"
 
Mass.Refresh
'*** ���� ��������� ������ ������ *************

'DataEnvironment1.��������_Z
'DataEnvironment1.��������_�
'DataEnvironment1.��������_Spisok_Dob
'DataEnvironment1.��������_����������_����������_��_�����
'DataEnvironment1.��������_����������_������_�����_��_���
'DataEnvironment1.��������_�������_��_�����
'DataEnvironment1.�������_����������_��_�����
'DataEnvironment1.�������_����������_���_�����_��_�����
'******** ����� ����� ��������� ������ ������ *******

' ************************ ���� ���������� ������ *******
'DataEnvironment1.����������_Z
'DataEnvironment1.����������_�
'DataEnvironment1.��������

'DataEnvironment1.�������_������_������_�����_��������_��
'DataEnvironment1.�����_�_����������_��_�����
'DataEnvironment1.�����_���_�����_��_�����
'DataEnvironment1.����������_����������_��_�����
'DataEnvironment1.�����_������_�����_��_�����
'DataEnvironment1.�����_�������_��_�����
' ************************ ����� ����� ���������� ������ *******


'DataEnvironment1.Init_Zall


Mass.Hide

End Sub

Private Sub Dir1_Change()

'Dir1.Path = Drive1.Drive
'Dir1.Refresh
Dir1.Refresh
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Dir1.Refresh
End Sub



Private Sub OKButton_Click()
Dialog.Hide
Call Form1.MakePt(Pt)
'Form1.Refresh
Form1.Show
End Sub
