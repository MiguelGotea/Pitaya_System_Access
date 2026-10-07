' ==========================================================
' Modulo  : Form_Resumen Pagos Proovedor
' Tipo    : 100
' Lineas  : 147
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database
Sub guardarresumen()
'guardar recumen de pago en imagen

Set Application.Printer = Application.Printers(elegirimpresora("ImpresoraConvertirImagen"))

[Forms]![Resumen Pagos Proovedor].Form.SetFocus
[Forms]![Resumen Pagos Proovedor].Form.Requery
[Forms]![Resumen Pagos Proovedor].Form.Printer.Orientation = acPRORPortrait
[Forms]![Resumen Pagos Proovedor].Form.Printer.PaperSize = acPRPSLetter
[Forms]![Resumen Pagos Proovedor].Form.Printer.LeftMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.Printer.RightMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.Printer.BottomMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.Printer.TopMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.SetFocus
DoCmd.PrintOut , 1, 1

Set Application.Printer = Application.Printers(elegirimpresora("ImpresoraTintaOficinas"))

'Carpeta para archivo
Dim provee As String

provee = DLookup("[Nombre]", "[Proovedores]", "[CodProovedor]=" & Me.CodigoBusqueda)
Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Pago Proovedores\" & provee & "\" & nombreciudadsucursalglobal(Me.sucursalbusqueda)
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Pago Proovedores\" & provee
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

Sleep (5000)

'copiar archivo a otra carpeta
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\imagentemporal.png"
h = Direccion & "\" & Me.aresumen & ".png"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

End Sub

Private Sub Comando209_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO subresumenpago(codresumenpago, monto)" & _
" values (" & Me.aresumen & ", 0)"
DoCmd.SetWarnings True

Me.Requery
End Sub

Private Sub Comando2130_Click()
[Forms]![Resumen Pagos Proovedor].Form.SetFocus
[Forms]![Resumen Pagos Proovedor].Form.Requery
[Forms]![Resumen Pagos Proovedor].Form.Printer.Orientation = acPRORLandscape
[Forms]![Resumen Pagos Proovedor].Form.Printer.PaperSize = acPRPSLetter
[Forms]![Resumen Pagos Proovedor].Form.Printer.LeftMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.Printer.RightMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.Printer.BottomMargin = 0
[Forms]![Resumen Pagos Proovedor].Form.Printer.TopMargin = 0
DoCmd.PrintOut , 1, 1

Call guardarresumen

End Sub

Private Sub Comando289_Click()
Dim provee As String
Dim direc As String

provee = DLookup("[Nombre]", "[Proovedores]", "[CodProovedor]=" & Me.CodigoBusqueda)
direc = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Pago Proovedores\" & provee & "\" & nombreciudadsucursalglobal(Me.sucursalbusqueda)

Call guardarresumen

'enviar arcivo por telegram

Call EnviarImagenTelegram(direc & "\", Me.aresumen & ".png", provee & " , " & nombreciudadsucursalglobal(Me.sucursalbusqueda) & " , " & Me.TotalTotal, grupotcontabilidad())

MsgBox "Pago enviado"
End Sub

Private Sub Comando304_Click()
Dim provee As String
Dim direc As String
Dim numtel As String
Dim mensa As String
Dim archivo As String

provee = DLookup("[Nombre]", "[Proovedores]", "[CodProovedor]=" & Me.CodigoBusqueda)
direc = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Pago Proovedores\" & provee & "\" & nombreciudadsucursalglobal(Me.sucursalbusqueda)
numtel = DLookup("[Numero]", "[Proovedores]", "[CodProovedor]=" & Me.CodigoBusqueda)

Call guardarresumen

'enviar archivo por wsp

mensa = primersaludo() & ", verificar el resumen de facturas de la semana "
archivo = direc & "\" & Me.aresumen & ".png"

Call EnviarImagenWhatsapp(mensa, archivo, numtel)

MsgBox "Mensaje enviado"
End Sub

Private Sub Comando820_Click()
If MsgBox("¿Estas seguro de querer eliminar el registro?.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM subresumenpago WHERE codsubresumenpago = " & Me.codsubresumenpago
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.sucursalbusqueda.RowSource = "SELECT DatosSistema.Ciudad, DatosSistema.CodSistema FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'"

Me.InsideHeight = 14000
End Sub

Private Sub CodigoBusqueda_Exit(Cancel As Integer)
Me.Requery
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE resumenpago SET resumenpago.CodProovedor = " & Me.CodigoBusqueda & _
" WHERE resumenpago.codresumenpago = " & Me.aresumen
DoCmd.SetWarnings True
End Sub


Private Sub monto_Exit(Cancel As Integer)
Me.Requery
Me.fechapago.SetFocus
End Sub

Private Sub sucursalbusqueda_Exit(Cancel As Integer)
Me.Requery
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE resumenpago SET resumenpago.sucursal = " & Me.sucursalbusqueda & _
" WHERE resumenpago.codresumenpago = " & Me.aresumen
DoCmd.SetWarnings True
End Sub
