VERSION 5.00
Begin VB.Form MenuZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Меню"
   ClientHeight    =   3372
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   2808
   Icon            =   "MenuZ.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3372
   ScaleWidth      =   2808
   StartUpPosition =   2  'CenterScreen
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "Из банка DBF формат"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   492
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "ТХТ формат одна услуга"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "Выход"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton4 
      Height          =   492
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "ТХТ формат несколько услуг"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton5 
      Height          =   492
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "Импорт из ЕРКЦ "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton6 
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "Импорт из соцгарантий"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton7 
      Height          =   492
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      Caption         =   "Почта XLS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MenuZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Unload Me
MainMenu.Enabled = True
End Sub

Private Sub xpcmdbutton1_Click()
BankTXTimpott.Show 1
Unload Me
End Sub

Private Sub xpcmdbutton2_Click()
BankImport.Show 1
Unload Me
End Sub

Private Sub xpcmdbutton3_Click()
Unload Me
End Sub



Private Sub xpcmdbutton4_Click()
BankTXTimpott_91.Show 1
Unload Me
End Sub

Private Sub xpcmdbutton5_Click()
ERKCBankImpor.Show 1
Unload Me
End Sub

Private Sub xpcmdbutton6_Click()
BankSocGarimpott.Show
Unload Me
End Sub

Private Sub xpcmdbutton7_Click()
PoctaLift.Show 1
Unload Me
End Sub
