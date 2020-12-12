VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form PrintW 
   Caption         =   "PrintW"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   LinkTopic       =   "PrintW"
   ScaleHeight     =   8595
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VSPrinter8LibCtl.VSPrinter VP 
      Align           =   3  'Align Left
      Height          =   8595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _cx             =   20346
      _cy             =   15161
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   1
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   500
      MarginTop       =   200
      MarginRight     =   500
      MarginBottom    =   200
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   46.2154942119323
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   -1  'True
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
End
Attribute VB_Name = "PrintW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 



Private Sub Form_Load()
VP.PrintDialog pdPrinterSetup
'VP.ShowGuides

End Sub
 'BeforePageBreak: controls page breaks

    ' мы принимаем что у нас есть промежуточные суммы выше детали,

    ' и предохран€ть колонки промежуточной суммы от ' последний на странице
    Private Sub fg_BeforePageBreak(ByVal Row As Long, BreakOK As Boolean)

        ' если эта колонка €вл€етс€ промежуточной суммой, возглавл€ющей, мы не можем ломать здесь
        If FG.IsSubtotal(Row) Then BreakOK = False
        
    End Sub


    ' GetHeaderRow: ѕќ—“ј¬Ћя≈“ заголовку колонки дл€ новых страниц

    ' мы принимаем что у нас есть колонки названи€
    'с RowData установленное, чтобы -1 ', которые мы хотим показать

    ' выше данных
    Private Sub fg_GetHeaderRow(ByVal Row As Long, HeaderRow As Long)

        Dim R As Long

    

       ' игнорироватьс€ если верхн€€ колонка €вл€етс€ заголовком уже
        If FG.RowData(Row) = -1 Then Exit Sub
    
        ' нам нужно заголовок, так что находить один
        For R = FG.FixedRows To fa.Rows - 1
            If FG.RowData(R) = -1 Then
                HeaderRow = R
                Exit Sub
            End If
        Next
        
        
        
        
    End Sub



