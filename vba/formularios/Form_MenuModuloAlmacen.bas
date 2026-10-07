' ==========================================================
' Modulo  : Form_MenuModuloAlmacen
' Tipo    : 100
' Lineas  : 84
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database


Private Sub Comando1101_Click()
DoCmd.OpenForm "Control Semanal Main"

End Sub

Private Sub Comando1217_Click()
DoCmd.Quit
End Sub




Private Sub Comando1311_Click()
DoCmd.OpenForm "ControlInventarioPorciones"

End Sub

Private Sub Comando1674_Click()
Call eliminartablasmain

Call importartablasmain

MsgBox "Tablas Main Actualizadas"
End Sub


Private Sub Comando1702_Click()
Call descargardatoscentralcopia("almacen")
MsgBox "Datos de la central completos"
End Sub

Private Sub Comando2022_Click()
Call importartablaespecifica("Contabilidad", "Compras", "Compras", 1)
DoCmd.OpenForm "Corroborar Ingresos Pitaya"
End Sub





Private Sub Comando2037_Click()
If MsgBox("El procedimiento correcto es: " & Chr(13) & Chr(10) & "1. Cerrar sistema" & Chr(13) & Chr(10) & "2. Entrar a IMPRESORAS Y DISPOSITIVOS y elegir a XPRINTER como impresora predefinida" & Chr(13) & Chr(10) & "3. Abrir nuevamente el sistema y entrar a la herramientas de imprimir codigo" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Tiene seleccionado la impresora XPRINTER como PREDETERMINADO y reinicado el sistema?", vbYesNo, "CONFIRMACION") = vbYes Then
    Dim cluba As Long
    cluba = InputBox("ingrese codigo de la membresia", "CODIGO")
    
    DoCmd.OpenReport "StickerClientesClub", acViewReport
    [Reports]![StickerClientesClub]![mcodigo] = cluba
    DoCmd.SelectObject acReport, "StickerClientesClub"
    DoCmd.PrintOut acSelection, 1, 1
    DoCmd.Close acReport, "StickerClientesClub"
End If
End Sub

Private Sub Comando2040_Click()
DoCmd.OpenForm "Control Semanal No Variables"
End Sub

Private Sub Comando2050_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "Almacen"
End Sub

Private Sub Comando2090_Click()
DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").modocentral
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "BATIDOS PITAYA 0"
Me.ShortcutMenu = False

End Sub

Private Sub Comando1996_Click()
Call importartablaespecifica("Contabilidad", "Compras", "Compras", 1)
DoCmd.OpenForm "GenerarIngresoDeCompras"
End Sub



