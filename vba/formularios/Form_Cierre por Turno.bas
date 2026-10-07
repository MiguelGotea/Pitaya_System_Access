' ==========================================================
' Modulo  : Form_Cierre por Turno
' Tipo    : 100
' Lineas  : 61
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database

Private Sub Comando222_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]<>0 AND [Transferencia]=0"
End Sub

Private Sub Comando107_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "#"

End Sub

Private Sub Comando221_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]=0 AND [Transferencia]=0"
End Sub

Private Sub Comando284_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Me.afecha & "# AND [POS]<>0 AND [Comision]=0 AND [Transferencia]<>0"
End Sub

Public Sub Comando419_Click()



'Nuevo preingreso
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO CierreDiario(HoraInicial, HoraFinal, Fecha, CodOperario, TotalPedidosYa, TotalTransferencia, TotalPOS)" & _
" values (#" & Format(Me.aHoraInicial, "hh:nn:ss") & "#, #" & Format(Me.aHoraFinal, "hh:nn:ss") & "#, #" & Me.afecha & "#, " & Me.aCodOperario & ", " & Me.aTotalPedidosYa & ", " & Me.aTotalTransferencia & ", " & Me.aTotalPOS & ")"
DoCmd.SetWarnings True

Dim cierrel As Long
cierrel = CierreFinal(Date)




' Abrir formulario "Contador Caja" en vista de formulario
DoCmd.OpenForm "Contador Caja", acNormal, , "[CodigoCierre]=" & cierrel
[Forms]![Contador Caja]![titulo1] = "CONTEO DE CIERRE"
[Forms]![Contador Caja]![totalentregarsistema] = [Forms]![Cierre por Turno]![cajafinal]
'DoCmd.OpenForm "Contador Caja", acNormal, , "CierreFinal(" & Forms("Cierre por Turno")("afecha") & ")"
'DoCmd.Close acForm, "Cierre por Turno"

[Forms]![Main Pitaya].Form.Requery

End Sub

Private Sub Comando426_Click()
DoCmd.Close acForm, "Cierre por Turno"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

Me.Comando107.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando221.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando222.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"
Me.Comando284.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Lupa.jpg"

End Sub

