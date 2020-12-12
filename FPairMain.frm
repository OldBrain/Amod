VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Fmain 
   Caption         =   "Main Form"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7695
      _cx             =   13573
      _cy             =   9551
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FPairMain.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   10
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show SubForm"
      Height          =   195
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2775
   End
End
Attribute VB_Name = "Fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' *************************************************************************
'  Copyright ©1999 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit
Implements IMessageSink

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5
Private Const WM_SIZING = &H214
Private Const WM_MOVING = &H216

Private Const WM_ACTIVATE = &H6
Private Const WM_NCACTIVATE = &H86

Private WithEvents m_Sub As FPair
Attribute m_Sub.VB_VarHelpID = -1
Dim rsLg As ADODB.Recordset

Private Sub Check1_Click()
   If Check1.Value Then
      m_Sub.Show , Me
      Check1.Caption = "Hide Sub Form"
   Else
      m_Sub.Hide
      Check1.Caption = "Show Sub Form"
   End If
End Sub



Private Sub Form_Load()

Set rsLg = New ADODB.Recordset

rsLg.Open ("SELECT KLS_PRIV.N_KLS as  Œƒ, KLS_PRIV.NAME_KLS as À¸„ÓÚ‡ FROM KLS_PRIV order by name_kls"), Mconn

   ' Hook into messages for this window.
   Call HookWindow(Me.hWnd, Me)
   
   ' Create, load, and hook messages for sub form.
   Set m_Sub = New FPair
   Load m_Sub
   Call HookWindow(m_Sub.hWnd, Me)
   m_Sub.Show , Me
   
   
   
  Set FG.DataSource = rsLg
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Always unhook before unloading!
   Call UnhookWindow(m_Sub.hWnd)
   Unload m_Sub
   Set m_Sub = Nothing
   
   ' Unhook this (main) one too!
   Call UnhookWindow(Me.hWnd)
   Potok.Show
End Sub

Private Function IMessageSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   Static rMain As RECT
   Static rSub As RECT
   Dim Result As Long
   
   Select Case hWnd
      Case Me.hWnd
         Select Case msg
            Case WM_MOVE, WM_MOVING, WM_SIZE
               ' Move subform to desired position.
               If Not (m_Sub Is Nothing) Then
                  ' Retrieve coordinates for both windows.
                  Call GetWindowRect(hWnd, rMain)
                  Call GetWindowRect(m_Sub.hWnd, rSub)
                  ' Position subform appropriately.
                  Call MoveWindow(m_Sub.hWnd, rMain.Right, rMain.Top, rSub.Right - rSub.Left, rMain.Bottom - rMain.Top, True)
                  ' Store new position of subform.
                  Call GetWindowRect(m_Sub.hWnd, rSub)
               End If
               Result = InvokeWindowProc(hWnd, msg, wp, lp)
               
            Case Else
               ' Pass along to default window procedure.
               Result = InvokeWindowProc(hWnd, msg, wp, lp)
         End Select
      
      Case m_Sub.hWnd
         Select Case msg
            Case WM_ACTIVATE
               ' Have main form retain active titlebar.
               Result = InvokeWindowProc(hWnd, msg, wp, lp)
               Call SendMessage(Me.hWnd, WM_NCACTIVATE, 1, ByVal 0&)
               
            Case WM_MOVING
               ' Copy stored position of subform to the position
               ' the user is trying to drag it to.
               Call CopyMemory(ByVal lp, rSub, Len(rSub))
               Result = 1
               
            Case Else
               ' Pass along to default window procedure.
               Result = InvokeWindowProc(hWnd, msg, wp, lp)
         End Select
      
      End Select
   ' Return desired result code to Windows.
   IMessageSink_WindowProc = Result
End Function

Private Sub m_Sub_Hide()
   ' User clicked [X] on subform.
   Check1.Value = False
End Sub
