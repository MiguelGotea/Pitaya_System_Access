' ==========================================================
' Modulo  : Form_CambiosGlobalRegistrados
' Tipo    : 100  |  Lineas: 52
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database



Private Sub Comando1224_Click()
Dim desti As String
Dim locaf As Integer

desti = DLookup("[Destino]", "[PreIngresoPitaya]", "[CodPreingresoPitaya]=" & Me.CodPreIngresoPitaya)
locaf = CInt(Split(desti, " ")(1))

    
DoCmd.OpenForm "RegistroPreIngresosPitaya", , , "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
               " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
[Forms]![RegistroPreingresosPitaya]![aencabezado].Caption = "DETALLE DE PREINGRESO " & UCase(desti) & " " & UCase(nombreciudadsucursalglobal(locaf))
[Forms]![RegistroPreingresosPitaya]![fechaac] = Me.Fecha
[Forms]![RegistroPreingresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
[Forms]![RegistroPreingresosPitaya]![codigoope] = 5
[Forms]![RegistroPreingresosPitaya]![validado] = IIf(vali = 0, "EN REVISION", "VALIDADO")

[Forms]![RegistroPreingresosPitaya]![Comando633].Visible = IIf(statuscambiospreingresoporsucursal(Me.CodPreIngresoPitaya) = 0, True, False)
[Forms]![RegistroPreingresosPitaya]![Etiqueta668].Visible = True
[Forms]![RegistroPreingresosPitaya]![titulostatuscambios].Visible = True
[Forms]![RegistroPreingresosPitaya]![statuscambios] = statuscambiospreingresoporsucursal(Me.CodPreIngresoPitaya)

Call Forms("[RegistroPreIngresosPitaya]").modovistacentral

End Sub

Private Sub Comando155_Click()

Me.Requery

End Sub



Private Sub Form_Close()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM CambiosPreingresoSucursal"
DoCmd.SetWarnings True
End Sub

Private Sub Form_Open(Cancel As Integer)


Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
End Sub



