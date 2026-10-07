' ==========================================================
' Modulo  : Form_Inicio Modo Aplicativo
' Tipo    : 100
' Lineas  : 25
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database
Option Explicit

Const SW_HIDE = 0
Const SW_NORMAL = 1
Const SW_MINIMIZED = 2
Const SW_MAXIMIZED = 3




Private Sub Form_Open(Cancel As Integer)

Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Call ShowWindow(hWndAccessApp, SW_HIDE)
DoCmd.OpenForm "Main Pitaya", windowmode:=acDialog 'acDialog 'acWindowNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim lngRetCode As Long
lngRetCode = ShowWindow(hWndAccessApp, SW_MAXIMIZED)
End Sub


