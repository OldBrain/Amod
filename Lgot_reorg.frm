VERSION 5.00
Begin VB.Form Lgot_reorg 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5175
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6555
   ControlBox      =   0   'False
   Icon            =   "Lgot_reorg.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Начать реорганизацию"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   6015
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Text            =   "Новая льгота"
      Top             =   2280
      Width           =   5775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   480
      TabIndex        =   3
      Text            =   "Старая льгота"
      Top             =   960
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Новая льгота"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Image imgTitleHelp 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      Height          =   240
      Left            =   0
      Picture         =   "Lgot_reorg.frx":030A
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   5850
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Старая льгота"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   480
      Picture         =   "Lgot_reorg.frx":084C
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   720
      Picture         =   "Lgot_reorg.frx":0F96
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   240
      Picture         =   "Lgot_reorg.frx":16E0
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Lgot_reorg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLg As ADODB.Recordset
Dim oldN As String
Dim NewN As String

Private Sub Combo1_LostFocus()
oldN = Val(Mid(Combo1.Text, 1, InStr(2, Combo1.Text, " ")))


End Sub

Private Sub Combo2_LostFocus()
NewN = Val(Mid(Combo2.Text, 1, InStr(2, Combo2.Text, " ")))
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

If Combo1.Text = "Старая льгота" Then
Combo1.SetFocus
Exit Sub
End If
If Combo2.Text = "Новая льгота" Then
Combo2.SetFocus
Exit Sub
End If
If Combo1.Text = Combo2.Text Then
MsgBox "Зачем реорганизовывать льготу саму в себя?"
Exit Sub
End If



If MsgBox("Провести реорганизазию льготы " + vbNewLine + Combo1.Text + vbNewLine + " на льготу " + Combo2.Text, vbYesNo) = vbYes Then


Pod.Show
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = 1000
Pod.ProgressBar1.Value = 1
Mconn.Execute ("UPDATE tmp_lgota SET tmp_lgota.KodKls = " + NewN + " WHERE (((tmp_lgota.KodKls)=" + oldN + "))")
Pod.ProgressBar1.Value = 500
Mconn.Execute ("UPDATE Lgota SET Lgota.Numer = " + NewN + " WHERE (((Lgota.Numer)=" + oldN + "))")
Pod.ProgressBar1.Value = 700
Mconn.Execute ("UPDATE OtheOwner SET OtheOwner.PRIVILEGE = " + NewN + " WHERE (((OtheOwner.PRIVILEGE)=" + oldN + "))")
Pod.ProgressBar1.Value = 1000
Unload Pod
msg.Show

msg.Label1.Caption = "Реорганизазия льготы проведена успешно!" + vbNewLine + "Теперь льгота " + vbNewLine + Combo1.Text + vbNewLine + " переименована в " + vbNewLine + Combo2.Text + vbNewLine + ". Льготу " + Combo1.Text + vbNewLine + " можно удалить." + vbNewLine + vbNewLine + "Не забудьте пересчитать лицевые счета."
Unload Me
End If
End Sub

Private Sub Form_Load()
lblTitle = "Реорганизация льгот"
MakeWindow Me, True

Set rsLg = New ADODB.Recordset

rsLg.Open ("SELECT KLS_PRIV.N_KLS, KLS_PRIV.NAME_KLS FROM KLS_PRIV order by KLS_PRIV.N_KLS"), Mconn
rsLg.MoveFirst
Do While Not rsLg.EOF
Combo1.AddItem Str(rsLg("n_kls")) + "  " + rsLg("name_kls")
Combo2.AddItem Str(rsLg("n_kls")) + "  " + rsLg("name_kls")
rsLg.MoveNext
Loop
rsLg.Close
End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub
