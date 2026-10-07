' ==========================================================
' Modulo  : Form_Control Diario Cierres y Compras
' Tipo    : 100
' Lineas  : 111
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database

Private Sub Comando33_Click()
Me.Requery
End Sub





Private Sub Comando496_Click()
Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & numerosemana(Me.dfecha)
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

Dim strDefaultPrinter As String
strDefaultPrinter = Application.Printer.DeviceName
archivo = Direccion & "\" & numerosemana(Me.dfecha) & " - Reporte Cierre Total.pdf"
Set Application.Printer = Application.Printers("Microsoft XPS Document Writer")
DoCmd.OutputTo acOutputReport, "ReporteCierresSemana", acFormatPDF, archivo, False
Set Application.Printer = Application.Printers(strDefaultPrinter)
End Sub



Private Sub Comando540_Click()
DoCmd.OpenForm "Busqueda Compras", acNormal, , "[Fecha]=#" & Me.dfecha & "#"
End Sub

Private Sub Comando107_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.dfecha & "#"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Me.dfecha) & "/" & Day(Me.dfecha) & "/" & Year(Me.dfecha) & " " & " TOTALES"
[Forms]![HistorialVentasFiltro].tipofiltro = "TODO"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False
End Sub

Private Sub Comando530_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.dfecha & "# AND [POS]<>0 AND ([Delivery]=5 Or [Delivery]=6 Or [Delivery]=8)"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Me.dfecha) & "/" & Day(Me.dfecha) & "/" & Year(Me.dfecha) & " " & " PEDIDOS YA"
[Forms]![HistorialVentasFiltro].tipofiltro = "PEDIDOSYA"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False
End Sub

Private Sub Comando557_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.dfecha & "# AND [POS]<>0 AND [Delivery]=0 AND [Transferencia]<>0"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Me.dfecha) & "/" & Day(Me.dfecha) & "/" & Year(Me.dfecha) & " " & " TRANSFERENCIA"
[Forms]![HistorialVentasFiltro].tipofiltro = "TRANSFERENCIA"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False
End Sub

Private Sub Comando1083_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.dfecha & "# AND [POS]=0 AND [Comision]=0 AND [Transferencia]=0"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Me.dfecha) & "/" & Day(Me.dfecha) & "/" & Year(Me.dfecha) & " " & " EFECTIVO"
[Forms]![HistorialVentasFiltro].tipofiltro = "EFECTIVO"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False
End Sub

Private Sub Comando90_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.dfecha & "# AND [POS]<>0 AND [Delivery]=0 AND [Transferencia]=0"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Me.dfecha) & "/" & Day(Me.dfecha) & "/" & Year(Me.dfecha) & " " & " POS"
[Forms]![HistorialVentasFiltro].tipofiltro = "POS"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False
End Sub

Private Sub Comando787_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Comando540.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando107.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando90.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando530.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando557.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando1083.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
End Sub
