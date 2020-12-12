VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Dim AboutBox As New AboutBox
With AboutBox
    .Title = "Some Application"
    .Version = "Version 1.2.3.4"
    .Company = "Some Company (R)"
    .Copyright = "(C) Some Company 1900-2004"
    .Description = "The quick brown fox jumps over a lazy dog again and again"
    .License = "This sample is hosted at <A HREF=""http://vbrussian.com/Examples.asp?ID=100"">vbrussian.com</A>. Visit it for more info!"
    .hWndOwner = Me.hWnd
    Set .Icon = Me.Icon
    .AboutBox
End With
End Sub
