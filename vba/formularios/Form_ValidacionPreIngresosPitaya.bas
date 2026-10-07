' ==========================================================
' Modulo  : Form_ValidacionPreIngresosPitaya
' Tipo    : 100  |  Lineas: 46
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database




Private Sub Comando214_Click()
Me.Requery
If Me.validado = "VALIDADO" Then

    'If ContadorItemsPreingreso(Me.CodPreIngresoPitaya) > 30 Then
    'MsgBox "La lista tiene mas de 30 items, saldra una hoja adicional", vbCritical
    'End If
    
    'Me.tipoimpresion = 2 'COPIA DESPACHO
    'DoCmd.OpenReport "DetallePreingreso2", acViewNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    
    Me.tipoimpresion = 3 'COPIA ALMACEN
    'DoCmd.OpenReport "DetallePreingreso2", acViewNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    DoCmd.OpenForm "DetallePreingreso2", acNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    '[Forms]![DetallePreingreso2].Form.Requery
    DoCmd.PrintOut
    DoCmd.Close acForm, "DetallePreingreso2"
    
    Me.tipoimpresion = 4 'COPIA SUCURSAL
    'DoCmd.OpenReport "DetallePreingreso2", acViewNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    DoCmd.OpenForm "DetallePreingreso2", acNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    '[Forms]![DetallePreingreso2].Form.Requery
    DoCmd.PrintOut
    DoCmd.Close acForm, "DetallePreingreso2"
Else
    MsgBox "Pendiente de validacion para poder imprimir la lista, se imprimira solo una copia no valida para despacho"
    Me.tipoimpresion = 1 'PRELIMINAR NO VALIDO PARA DESPACHO
    'DoCmd.OpenReport "DetallePreingreso2", acViewNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    DoCmd.OpenForm "DetallePreingreso2", acNormal, , "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & " AND PorcionDentroDeMezcla([CodCotizacion])=0"
    '[Forms]![DetallePreingreso2].Form.Requery
    DoCmd.PrintOut
    DoCmd.Close acForm, "DetallePreingreso2"
End If
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

