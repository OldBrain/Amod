VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOstrovodrepInhs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Подготовка карточки квартиросъемщика"
   ClientHeight    =   7380
   ClientLeft      =   2436
   ClientTop       =   2040
   ClientWidth     =   8772
   Icon            =   "frmOstrovodrepInhs.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8772
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Удалить"
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      ToolTipText     =   "Ctrl+-"
      Top             =   6930
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   4050
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Добавить"
      Height          =   375
      Left            =   2430
      TabIndex        =   2
      ToolTipText     =   "Ctrl++"
      Top             =   6915
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   6915
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отчет"
      Default         =   -1  'True
      Height          =   375
      Left            =   5610
      TabIndex        =   4
      Top             =   6915
      Width           =   1485
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   1845
      Left            =   270
      TabIndex        =   1
      Top             =   4920
      Width           =   7845
      _cx             =   13838
      _cy             =   3254
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmOstrovodrepInhs.frx":011A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2376
      Left            =   1920
      Picture         =   "frmOstrovodrepInhs.frx":01D2
      Top             =   216
      Width           =   3912
   End
   Begin VB.Label Label2 
      Caption         =   "Прописанные:"
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   4560
      Width           =   2385
   End
   Begin VB.Label Label1 
      Caption         =   "Ордер"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4110
      Width           =   645
   End
End
Attribute VB_Name = "frmOstrovodrepInhs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public frmowner As Form

Dim RSInputData As ADODB.Recordset



Private Sub Command1_Click()
    
    Dim i As Integer
    Dim nNum As Integer
    Dim StrBuf As String
    Dim RSetLoc As ADODB.Recordset
    
    Dim nPATHsiz As Long
    Dim szBuffer As String
    
    Set RSetLoc = New ADODB.Recordset
    
    
    RSetLoc.Open "SELECT * FROM Settings INNER JOIN AstrRay ON Settings.Ray = AstrRay.KodAstrRay;", Mconn, adOpenForwardOnly, adLockReadOnly
    
    With RSInputData
        odrin.entitycardnum = IIf(.Fields("BanKN").Value = Null, Empty, .Fields("BanKN").Value)
        StrBuf = IIf(.Fields("FAM").Value = Null, Empty, .Fields("FAM").Value) & " " _
                & IIf(.Fields("IM").Value = Null, Empty, .Fields("IM").Value) & " " _
                & IIf(.Fields("OT").Value = Null, Empty, .Fields("OT").Value)
        odrin.entityfio = IIf(StrBuf = "", Empty, StrBuf)
        odrin.street = IIf(.Fields("NAIM_KLS").Value = Null Or .Fields("NAIM_KLS").Value = "", Empty, .Fields("NAIM_KLS").Value)
        odrin.house = IIf(.Fields("Num").Value = Null Or .Fields("Num").Value = "", Empty, .Fields("Num").Value)
        odrin.flat = IIf(.Fields("Num").Value = Null Or .Fields("Num").Value = "", Empty, .Fields("kv_num").Value)
        odrin.area = IIf(.Fields("ComSpace").Value = Null, 0, Val(.Fields("ComSpace").Value))
        odrin.orgname = IIf(RSetLoc.Fields("Settings.Name").Value = Null, "", RSetLoc.Fields("Settings.Name").Value)
        odrin.regionname = IIf(RSetLoc.Fields("AstrRay.Name").Value = Null, "", RSetLoc.Fields("AstrRay.Name").Value)
    End With
    odrin.order = Text1.Text
    
    RSetLoc.Close
    
    
    nNum = VSFlexGrid1.Rows - 1
    ReDim inhs(nNum) As inhabitantsstruct
    
    deleteinhabitants 'Удаление предидущего ввода жильцов, иначе жильци этого адреса выведутся с жильцами предидущего выводимого адреса
    
    For i = 1 To nNum
        With VSFlexGrid1
            inhs(i).FIO = IIf(.TextMatrix(i, 1) = "", Empty, .TextMatrix(i, 1))
            inhs(i).birthyear = Val(.TextMatrix(i, 2))
            inhs(i).relationship = IIf(.TextMatrix(i, 3) = "", Empty, .TextMatrix(i, 3))
            
            
            inhs(i).datain.wyear = IIf(.TextMatrix(i, 0) = "", 0, Year(CDate(.TextMatrix(i, 0))))
            inhs(i).datain.wmonth = IIf(.TextMatrix(i, 0) = "", 0, Month(CDate(.TextMatrix(i, 0))))
            inhs(i).datain.wday = IIf(.TextMatrix(i, 0) = "", 0, Day(CDate(.TextMatrix(i, 0))))
        End With
        
        addinhabitant inhs(i)
        
    Next i
    
    nNum = 1000
    
makebuffer:
    szBuffer = String(nNum, 0)
    
    nPATHsiz = GetEnvironmentVariable("PATH", szBuffer, nNum)
    If (nPATHsiz = nNum) Then
        nNum = nNum + 1000
        szBuffer = ""
        GoTo makebuffer
    End If
    
    szBuffer = Left(szBuffer, nPATHsiz) & ";" & App.Path & "\Util"
    SetEnvironmentVariable "PATH", szBuffer
    
    ostrovodrep odrin, App.Path & "\Rep", frmowner.hwnd
    
    Set RSetLoc = Nothing
    
    Unload Me
End Sub

'With inhs(i)
'            MsgBox .FIO & vbNewLine & _
'                .birthyear & vbNewLine & _
'                .relationship & vbNewLine & _
'                vbNewLine & vbNewLine & _
'                .datain.wyear & vbNewLine & _
'                .datain.wmonth & vbNewLine & _
'                .datain.wday
'        End With

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    VSFlexGrid1.AddItem ""
    If VSFlexGrid1.Row < 1 Then VSFlexGrid1.Row = 1
    Command1.Enabled = VSFlexGrid1.Rows > 1
End Sub

Private Sub Command4_Click()
    'MsgBox "Не реализовано"
    If VSFlexGrid1.Row > 0 Then VSFlexGrid1.RemoveItem
    Command1.Enabled = VSFlexGrid1.Rows > 1
    'Caption = VSFlexGrid1.Row
End Sub

Private Sub Form_Load()

    Dim RSetLoc As ADODB.Recordset
    
    Set RSInputData = rsODRGlobal
    Set rsODRGlobal = Nothing
    
    With Screen
        Label1.Left = 4 * Screen.TwipsPerPixelX
        Label2.Left = 4 * Screen.TwipsPerPixelX
        Text1.Move Label1.Left + Label1.Width + 4 * Screen.TwipsPerPixelX _
        , Text1.Top, ScaleWidth - (Label1.Left + Label1.Width + 4 * Screen.TwipsPerPixelX) - 4 * Screen.TwipsPerPixelX
        
        
        VSFlexGrid1.Move .TwipsPerPixelX * 4, VSFlexGrid1.Top, Me.ScaleWidth - (.TwipsPerPixelX * 8)
        
        VSFlexGrid1.RowHeight(0) = VSFlexGrid1.RowHeight(0) * 2
        
        Image1.Left = ScaleWidth / 2 - Image1.Width / 2
        
        Command1.Top = ScaleHeight - Command1.Height - 4 * .TwipsPerPixelY
        Command2.Top = ScaleHeight - Command2.Height - 4 * .TwipsPerPixelY
        Command3.Top = ScaleHeight - Command3.Height - 4 * .TwipsPerPixelY
        Command4.Top = ScaleHeight - Command4.Height - 4 * .TwipsPerPixelY
    End With
    
    Set RSetLoc = New ADODB.Recordset
    
    
    
'    If F <> "" Then f1 = "WHERE (((OtheOwner.Numer)=" & F & "))"
'    sq = "SELECT OtheOwner.Numer, OtheOwner.Dom, OtheOwner.KV, OtheOwner.FAM, OtheOwner.IM, OtheOwner.OT, OtheOwner.PRIVILEGE, OtheOwner.BIRTHDAY, OtheOwner.NFAMILY, OtheOwner.PASSPORT, OtheOwner.LDATEBEG, OtheOwner.LDATEEND, OtheOwner.OhteCode From OtheOwner " & f1
    
    RSetLoc.Open "SELECT OtheOwner.Numer, OtheOwner.Dom, OtheOwner.KV, OtheOwner.FAM, OtheOwner.IM" _
    & ", OtheOwner.OT, OtheOwner.PRIVILEGE, OtheOwner.BIRTHDAY, OtheOwner.NFAMILY, " & _
    "OtheOwner.PASSPORT, OtheOwner.LDATEBEG, OtheOwner.LDATEEND, OtheOwner.OhteCode " & _
    "FROM OtheOwner WHERE (((OtheOwner.Numer)=" & _
    IIf(RSInputData.Fields("Numer").Value = Null, "", RSInputData.Fields("Numer").Value) _
    & "))", Mconn, adOpenForwardOnly, adLockReadOnly
    
    'RSInputData.Fields("OLDNUM").Value
    
    Do Until RSetLoc.EOF
        VSFlexGrid1.AddItem _
            RSetLoc.Fields("LDATEBEG") & vbTab _
            & RSetLoc.Fields("FAM") & " " _
            & RSetLoc.Fields("IM") & " " _
            & RSetLoc.Fields("OT") & vbTab _
            & RSetLoc.Fields("BIRTHDAY")
        RSetLoc.MoveNext
    Loop
    
    RSetLoc.Close
    Set RSetLoc = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'RSInputData.Close
    Set RSInputData = Nothing
End Sub

Private Sub VSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        If (Shift = 1) Then
            If VSFlexGrid1.Row = 1 _
            And VSFlexGrid1.Col = 0 Then
                Text1.SetFocus
            End If
        Else
            If VSFlexGrid1.Row = VSFlexGrid1.Rows - 1 _
            And VSFlexGrid1.Col = VSFlexGrid1.Cols - 1 Then
                Command3.SetFocus
            End If
        End If
    End If
    
    If Shift = 2 Then
        Select Case KeyCode
        Case 107
            Command3_Click
        Case 109
            Command4_Click
        Case 187
            Command3_Click
        Case 189
            Command4_Click
        End Select
        
    End If
    'Caption = "KeyCode == " & KeyCode & " , Shift == " & Shift
End Sub
