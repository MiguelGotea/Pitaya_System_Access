' ==========================================================
' Modulo  : CapturaPantallaFormulario
' Tipo    : 1  |  Lineas: 134
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database

''GUARDAR FORMULARIO EN PORTAPAPLES


Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr


Const CF_BITMAP As Long = 2
'CAPTURA DE PANTALLA FORMULARIO

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef pclsid As GUID) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As LongPtr, ByVal hpal As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As LongPtr, ByVal filename As LongPtr, ByRef clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type



Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Const SRCCOPY = &HCC0020

Public Sub CapturarVentanaComoPNG(iruta As String)
Dim hwnd As LongPtr
hwnd = GetForegroundWindow() ' Captura la ventana activa

Dim r As RECT
GetWindowRect hwnd, r
Dim width As Long: width = r.Right - r.Left
Dim height As Long: height = r.Bottom - r.Top

Dim hdcSrc As LongPtr: hdcSrc = GetDC(hwnd)
Dim hdcMem As LongPtr: hdcMem = CreateCompatibleDC(hdcSrc)
Dim hbm As LongPtr: hbm = CreateCompatibleBitmap(hdcSrc, width, height)
Call SelectObject(hdcMem, hbm)
Call BitBlt(hdcMem, 0, 0, width, height, hdcSrc, 0, 0, SRCCOPY)
Call ReleaseDC(hwnd, hdcSrc)

Dim token As LongPtr
Dim inputs As GdiplusStartupInput
inputs.GdiplusVersion = 1
Call GdiplusStartup(token, inputs, 0)

Dim img As LongPtr
Call GdipCreateBitmapFromHBITMAP(hbm, 0, img)

Dim clsidPng As GUID
CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), clsidPng

Dim path As String
path = iruta
Call GdipSaveImageToFile(img, StrPtr(path), clsidPng, 0)

Call GdipDisposeImage(img)
Call GdiplusShutdown(token)
Call DeleteObject(hbm)
Call DeleteDC(hdcMem)

End Sub




Public Sub CapturarFormularioAlPortapapeles(frm As String)
Dim hwnd As LongPtr
hwnd = Forms(frm).hwnd

Dim r As RECT
GetWindowRect hwnd, r

Dim w As Long, h As Long
w = r.Right - r.Left
h = r.Bottom - r.Top

Dim hdcSrc As LongPtr, hdcDest As LongPtr, hBitmap As LongPtr, hOld As LongPtr

hdcSrc = GetWindowDC(hwnd)
hdcDest = CreateCompatibleDC(hdcSrc)
hBitmap = CreateCompatibleBitmap(hdcSrc, w, h)
hOld = SelectObject(hdcDest, hBitmap)

BitBlt hdcDest, 0, 0, w, h, hdcSrc, 0, 0, SRCCOPY

Call SelectObject(hdcDest, hOld)
Call DeleteDC(hdcDest)
Call ReleaseDC(hwnd, hdcSrc)

' Copiar al portapapeles
OpenClipboard 0
EmptyClipboard
SetClipboardData CF_BITMAP, hBitmap
CloseClipboard

' Nota: No debes eliminar el bitmap si lo estás dejando en el portapapeles
'MsgBox "Formulario copiado al portapapeles como imagen.", vbInformation
End Sub

