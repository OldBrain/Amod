VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
''Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, ByVal lpTemplate As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Type DlgInfo
    Icon As Long
    App As Long
    OtherStuff As Long
End Type

Private m_Icon As StdPicture, m_hWndOwner As Long, m_Company As String, m_Title As String, m_Version As String
Private m_Copyright As String, m_Description As String, m_License As String
Private m_DlgProc As Long, m_IsNT As Boolean, hLib As Long
Private m_Redirector(0 To 7) As Long, m_CallAddress As Long

Public Property Set Icon(ByVal vData As StdPicture)
    Set m_Icon = vData
End Property

Public Property Get Icon() As StdPicture
    Set Icon = m_Icon
End Property

Public Property Let hWndOwner(ByVal vData As Long)
    m_hWndOwner = vData
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = m_hWndOwner
End Property


Public Property Let Company(ByVal vData As String)
    m_Company = vData
End Property

Public Property Get Company() As String
    Company = m_Company
End Property

Public Property Let Title(ByVal vData As String)
    m_Title = vData
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Version(ByVal vData As String)
    m_Version = vData
End Property

Public Property Get Version() As String
    Version = m_Version
End Property

Public Property Let Copyright(ByVal vData As String)
    m_Copyright = vData
End Property

Public Property Get Copyright() As String
    Copyright = m_Copyright
End Property

Public Property Let Description(ByVal vData As String)
    m_Description = vData
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let License(ByVal vData As String)
    m_License = vData
End Property

Public Property Get License() As String
    License = m_License
End Property

Private Sub Class_Initialize()
Dim lpShellAbout As Long, Signa As Long, osvi As OSVERSIONINFO
    hLib = LoadLibrary("shell32.dll")
    With osvi
        .dwOSVersionInfoSize = Len(osvi): GetVersionEx osvi
        m_IsNT = (.dwPlatformId = 2)
    End With
    If m_IsNT Then
        lpShellAbout = GetProcAddress(hLib, "ShellAboutW")
    Else
        lpShellAbout = GetProcAddress(hLib, "ShellAboutA")
    End If
    Do
        CopyLong lpShellAbout, Signa
        lpShellAbout = lpShellAbout + 1
    Loop Until Signa = &H381068
    CopyLong lpShellAbout - 8, m_DlgProc
    
    CopyLong lpShellAbout + 4, Signa
    m_CallAddress = IIf((Signa And &HFF) = &HFF, lpShellAbout + 10, lpShellAbout + 5)
    
    m_Redirector(0) = &HFFEC8B55 'push ebp: mov ebp, esp: push dword ptr
    m_Redirector(1) = &H75FF1475 '[ebp + 14h]: push dword ptr [ebp +
    m_Redirector(2) = &HC75FF10  '10h]:push dword ptr [ebp + 0ch]
    m_Redirector(3) = &H680875FF 'push dword ptr [ebp + 8]: push
    m_Redirector(4) = hLib
    m_Redirector(5) = &H68909090 'nop: nop: nop: push
    m_Redirector(6) = m_CallAddress
    m_Redirector(7) = &HCCCCCCC3 'ret: int 3: int 3: int 3

    m_Company = App.CompanyName
    m_Title = App.Title
    m_Version = App.Major & "." & App.Minor & "." & App.Revision
    m_Copyright = App.LegalCopyright
    m_Description = App.Comments
End Sub

Private Sub Class_Terminate()
    FreeLibrary hLib
    Unload About
End Sub

Public Sub AboutBox()
Dim DlgInfo As DlgInfo, App() As Byte, OtherStuff() As Byte
With DlgInfo
    If m_IsNT Then
        .App = StrPtr(m_Title)
        .OtherStuff = StrPtr(m_Description)
    Else
        App() = StrConv(m_Title & vbNullChar, vbFromUnicode)
        OtherStuff() = StrConv(m_Description & vbNullChar, vbFromUnicode)
        .App = VarPtr(App(0))
        .OtherStuff = VarPtr(OtherStuff(0))
    End If
    If m_Icon Is Nothing Then
        .Icon = 0
    Else
        .Icon = m_Icon
    End If
End With
    InitVariables m_Company, m_Copyright, m_Version, m_License, m_DlgProc
''    DialogBoxParam hLib, &H3810, m_hWndOwner, AddressOf DialogProc, VarPtr(DlgInfo)
    CallWindowProc VarPtr(m_Redirector(0)), &H3810, m_hWndOwner, AddressOf DialogProc, VarPtr(DlgInfo)
End Sub

