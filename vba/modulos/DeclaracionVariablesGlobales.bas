' ==========================================================
' Modulo  : DeclaracionVariablesGlobales
' Tipo    : 1  |  Lineas: 23
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long



Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As LongPtr, _
    ByVal dwFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000






