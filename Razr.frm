VERSION 5.00
Begin VB.Form Razr 
   Caption         =   "       Разрешение экрана"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   2535
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ок"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "1024 Х 768"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "800 Х 600"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      Caption         =   "Установите разрешение экрана"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Razr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const WM_DISPLAYCHANGE = &H7E
Const HWND_BROADCAST = &HFFFF&
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const BITSPIXEL = 12

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type



Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim OldX As Long, OldY As Long, nDC As Long, R As Integer

Private Sub Form_Unload(Cancel As Integer)
    'восстановите экранное решение
    ChangeRes OldX, OldY, GetDeviceCaps(nDC, BITSPIXEL)
    'удалите наш контекст устройства
    DeleteDC nDC
End Sub
Sub ChangeRes(X As Long, Y As Long, Bits As Long)
    Dim DevM As DEVMODE, ScInfo As Long, erg As Long, an As VbMsgBoxResult
    'Получите инфо в DevM
    erg = EnumDisplaySettings(0&, 0&, DevM)
    'Это - что мы собираемся изменяться
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = X 'ScreenWidth
    DevM.dmPelsHeight = Y 'ScreenHeight
    DevM.dmBitsPerPel = Bits '(can be 8, 16, 24, 32 or even 4)
    'Теперь измените показ и проверяйте если возможно
    erg = ChangeDisplaySettings(DevM, CDS_TEST)
    'Check if succesfull
    Select Case erg&
        Case DISP_CHANGE_RESTART
            an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
            If an = vbYes Then
                erg& = ExitWindowsEx(EWX_REBOOT, 0&)
            End If
        Case DISP_CHANGE_SUCCESSFUL
            erg = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
            ScInfo = Y * 2 ^ 16 + X
            'Уведомите все окно об экранном изменении решения
            SendMessage HWND_BROADCAST, WM_DISPLAYCHANGE, ByVal Bits, ByVal ScInfo
            MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
        Case Else
            MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
    End Select
End Sub

'111111111111111111111111111111111111111111111111111

Private Sub Command1_Click()
Dim nDC As Long
    'извлеките экранное решение
    OldX = Screen.Width / Screen.TwipsPerPixelX
    OldY = Screen.Height / Screen.TwipsPerPixelY
    'Создайте устройство контекстное, совместимое с экраном
    nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    
    If R = 1 Then
    'Измените экранное решение 800, 600
    ChangeRes 800, 600, GetDeviceCaps(nDC, BITSPIXEL)
   End If
   
    If R = 2 Then
    'Измените экранное решение 1024, 768
    ChangeRes 1024, 768, GetDeviceCaps(nDC, BITSPIXEL)
   End If
Razr.Hide
End Sub

Private Sub Option1_Click()
R = 1
End Sub

Private Sub Option2_Click()
R = 2
End Sub
