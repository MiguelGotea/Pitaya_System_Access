' ==========================================================
' Modulo  : Form_Ingreso de Compras
' Tipo    : 100  |  Lineas: 163
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database
Function statusbloqueofactura() As Integer

If (Me.Ingresado = True Or Me.Fecha <> Me.fechaac) And codigoLocal() <> 0 Then
    statusbloqueofactura = 0
Else
    statusbloqueofactura = 1
End If
End Function


Private Sub Comando112_Click()
If fechaac.Value < Date Then
    fechaac.Value = fechaac.Value + 1
    Me.Requery
End If
End Sub

Private Sub Comando113_Click()
fechaac.Value = fechaac.Value - 1

Me.Requery
End Sub

Private Sub Comando159_Click()
MsgBox "Lista de Tipos de Pago:" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
       "TIPO 1: Efectivo" & Chr(10) & Chr(13) & _
       "TIPO 3: Cuenta Pagos Corriente" & Chr(10) & Chr(13) & _
       "TIPO 4: Cuenta Depositos Dolares" & Chr(10) & Chr(13) & _
       "TIPO 5: Cuenta Depositos Cordobas" & Chr(10) & Chr(13) & _
       "TIPO 6: Credito Visa Connect Miles" & Chr(10) & Chr(13) & _
       "TIPO 7: Credito Ameican Express Pricesmart" & Chr(10) & Chr(13) & _
       "TIPO 8: Cuenta POS"
End Sub

Private Sub Comando163_Click()
MsgBox "Cantidades por unidad de presentacion: " & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
        "Galletas de Avena           " & Chr(9) & "Pricesmart" & Chr(9) & "120" & Chr(10) & Chr(13) & _
        "Chocolate Hersheys de 1.36kg" & Chr(9) & "Pricesmart" & Chr(9) & "2" & Chr(10) & Chr(13) & _
        "Azucar Golden Brown          " & Chr(9) & "Pricesmart" & Chr(9) & "3" & Chr(10) & Chr(13) & _
        "Caja Guantes                 " & Chr(9) & "Pricesmart" & Chr(9) & "2" & Chr(10) & Chr(13) & _
        "Cucharas                     " & Chr(9) & "Pricesmart" & Chr(9) & "250" & Chr(10) & Chr(13) & _
        "Barras Energeticas           " & Chr(9) & "Pricesmart" & Chr(9) & "30" & Chr(10) & Chr(13) & _
        "Granola Qoaker               " & Chr(9) & "Pricesmart" & Chr(9) & "2" & Chr(10) & Chr(13) & _
        "Bolsas para galletas         " & Chr(9) & "Otros         " & Chr(9) & "5"
End Sub

Private Sub Comando223_Click()
DoCmd.OpenForm "Lista de Proovedores"
End Sub

Private Sub Comando234_Click()
If codigoLocal() <> 0 And Me.fechaac <> Date Then
    MsgBox "no se pueden ingresar facturas de fechas distintas a la de hoy"

Else

    DoCmd.OpenForm "RegistroDeFacturaCompra"
    [Forms]![RegistroDeFacturaCompra]![afechaorden] = Me.fechaac
    [Forms]![RegistroDeFacturaCompra]![tipofacturaform] = "NUEVO"
    [Forms]![RegistroDeFacturaCompra]![Comando851].Enabled = False
    [Forms]![RegistroDeFacturaCompra]![condicion] = 1
    
    If Me.operarioorigen <> 0 Then
        [Forms]![RegistroDeFacturaCompra]![aoperario] = Me.operarioorigen
        [Forms]![RegistroDeFacturaCompra]![aoperario].Enabled = False
    End If
End If
End Sub

Private Sub Comando409_Click()

DoCmd.OpenForm "RegistroDeFacturaCompra"
[Forms]![RegistroDeFacturaCompra]![afechaorden] = Me.fechaac
[Forms]![RegistroDeFacturaCompra]![aproovedor] = Me.CodProveedor
[Forms]![RegistroDeFacturaCompra]![atipopago] = Me.Tipo
[Forms]![RegistroDeFacturaCompra]![aoperario] = Me.CodOperario
[Forms]![RegistroDeFacturaCompra]![afactura] = Me.NumeroFactura
[Forms]![RegistroDeFacturaCompra]![facturainvisible] = Me.NumeroFactura
[Forms]![RegistroDeFacturaCompra]![tipofacturaform] = "EDICION"
[Forms]![RegistroDeFacturaCompra]![condicion] = 1
[Forms]![RegistroDeFacturaCompra].Form.Requery

If statusbloqueofactura() = 0 Then
    MsgBox "Ya no se permiten ediciones de esta factura"
    Call Forms("RegistroDeFacturaCompra").bloquearedicionfactura
Else
    Call Forms("RegistroDeFacturaCompra").desbloquearedicionfactura
End If



End Sub

Private Sub EliminarFactura_Click()


If statusbloqueofactura() = 0 Then
    MsgBox "Esta factura ya esta bloqueada, no se puede eliminar"
Else
    If MsgBox("¿Estás seguro de querer eliminar la factura Nro " & NumeroFactura & " por un total de " & FacturaTotal & " C$? También se eliminarán los ingresos de productos.", vbYesNo + vbInformation, "Confirmar") = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE * FROM compras WHERE NumeroFactura = '" & NumeroFactura & "'"
        DoCmd.SetWarnings True
        Me.Requery
    End If
End If

End Sub

Private Sub dpproovedor_Exit(Cancel As Integer)
Me.Requery
End Sub

Private Sub dptipo_Exit(Cancel As Integer)
Me.Requery
End Sub






Private Sub fechaac_Change()

Me.Requery
End Sub



Private Sub Form_Close()
If CurrentProject.AllForms("Contador Caja").IsLoaded Then
    [Forms]![Contador Caja]![productoscomprassistema].height = 280 * CantidadProductosCompradosCajaDia(Date)
    [Forms]![Contador Caja].montocomprassistema.Requery
    [Forms]![Contador Caja].vvales.Requery
    [Forms]![Contador Caja].cantidadcomprassistema.Requery
    [Forms]![Contador Caja].productoscomprassistema.Requery
    [Forms]![Contador Caja].Form.Requery
End If

If CurrentProject.AllForms("Contador Caja 2").IsLoaded Then
    [Forms]![Contador Caja 2]![productoscomprassistema].height = 280 * CantidadProductosCompradosCajaDia(Date)

    [Forms]![Contador Caja 2].montocomprassistema.Requery
    [Forms]![Contador Caja 2].vvales.Requery
    [Forms]![Contador Caja 2].cantidadcomprassistema.Requery
    [Forms]![Contador Caja 2].productoscomprassistema.Requery
    [Forms]![Contador Caja 2].Form.Requery
End If

If CurrentProject.AllForms("Cierre por Turno 3").IsLoaded Then
    Call Forms("Cierre por Turno 3").actualizariconoscierre
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Comando409.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Editar.png"
Me.EliminarFactura.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Eliminar.png"

Me.Requery
End Sub

