' ==========================================================
' Modulo  : Form_Menu Modulo Contabilidad
' Tipo    : 100
' Lineas  : 77
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database


Private Sub Comando1702_Click()
Call descargardatoscentralcopia("contabilidad")
MsgBox "Datos de la central completos"
End Sub




Private Sub Comando285_Click()
DoCmd.Quit
End Sub

Private Sub Comando288_Click()
Call eliminartablasmain

Call importartablasmain

MsgBox "Tablas Main Actualizadas"
End Sub

Public Sub Comando2880_Click()
On Error GoTo Error

Dim Direccionano As String
Dim Direccionmes As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Year(Me.excelcomprasdesde)
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Year(Me.excelcomprasdesde) & "\" & Month(Me.excelcomprasdesde)
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If


DoCmd.OutputTo acOutputQuery, "HistorialComprasDatosCentral", acFormatXLSX, Direccionmes & "\" & codigoLocal() & " - Historial Compras Desde " & Format(Me.excelcomprasdesde, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.excelcomprashasta, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False
MsgBox "Datos Exportados en la carpeta del mes " & Month(Me.excelcomprasdesde)
Exit Sub
Error:
MsgBox "No se exporto el excel correctamente"
End Sub

Private Sub Comando311_Click()
DoCmd.OpenForm "HistorialOrdenDeCompra"
End Sub

Private Sub Comando312_Click()
DoCmd.OpenForm "HistorialResumenPagos"
End Sub




Private Sub Comando345_Click()
DoCmd.OpenForm "VerificarMarcacionesSemana"
End Sub

Private Sub Comando352_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "Contabilidad"
End Sub

Public Sub Comando85_Click()
DoCmd.OpenForm "Ingreso de Compras"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " ¦¦"
Me.ShortcutMenu = False
End Sub
