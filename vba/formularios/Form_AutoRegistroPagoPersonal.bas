' ==========================================================
' Modulo  : Form_AutoRegistroPagoPersonal
' Tipo    : 100
' Lineas  : 230
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database


Private Sub Comando39_Click()
'Linea 1
If Me.s1 <> 0 Then ' horas laboradas
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (104, 1, #" & Me.fech & "#, " & Me.s1 & ", '" & NombreOperario(Me.p1) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.ff1 <> 0 Then ' feriado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (306, " & cf1 & ", #" & Me.fech & "#, " & Me.ff1 & ", '" & NombreOperario(Me.p1) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If


If Me.v1 <> 0 Then ' vacaciones
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (196, " & hv1 & ", #" & Me.fech & "#, " & Me.v1 & ", '" & NombreOperario(Me.p1) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.b1 <> 0 Then 'hay bono
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (109, 1, #" & Me.fech & "#, " & Me.b1 & ", '" & NombreOperario(Me.p1) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.i1 <> 0 Then 'hay inss
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (413, 1, #" & Me.fech & "#, " & Me.i1 & ", '" & NombreOperario(Me.p1) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Linea 2
If Me.s2 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (104, 1, #" & Me.fech & "#, " & Me.s2 & ", '" & NombreOperario(Me.p2) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.ff2 <> 0 Then ' feriado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (306, " & CF2 & ", #" & Me.fech & "#, " & Me.ff2 & ", '" & NombreOperario(Me.p2) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.v2 <> 0 Then ' vacaciones
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (196, " & hv2 & ", #" & Me.fech & "#, " & Me.v2 & ", '" & NombreOperario(Me.p2) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If
    
If Me.b2 <> 0 Then 'hay bono
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (109, 1, #" & Me.fech & "#, " & Me.b2 & ", '" & NombreOperario(Me.p2) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.i2 <> 0 Then 'hay inss
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (413, 1, #" & Me.fech & "#, " & Me.i2 & ", '" & NombreOperario(Me.p2) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Linea 3
If Me.s3 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (104, 1, #" & Me.fech & "#, " & Me.s3 & ", '" & NombreOperario(Me.p3) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.ff3 <> 0 Then ' feriado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (306, " & cf3 & ", #" & Me.fech & "#, " & Me.ff3 & ", '" & NombreOperario(Me.p3) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.v3 <> 0 Then ' vacaciones
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (196, " & hv3 & ", #" & Me.fech & "#, " & Me.v3 & ", '" & NombreOperario(Me.p3) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If
    
If Me.b3 <> 0 Then 'hay bono
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (109, 1, #" & Me.fech & "#, " & Me.b3 & ", '" & NombreOperario(Me.p3) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.i3 <> 0 Then 'hay inss
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (413, 1, #" & Me.fech & "#, " & Me.i3 & ", '" & NombreOperario(Me.p3) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Linea 4
If Me.s4 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (104, 1, #" & Me.fech & "#, " & Me.s4 & ", '" & NombreOperario(Me.p4) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.ff4 <> 0 Then ' feriado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (306, " & cf4 & ", #" & Me.fech & "#, " & Me.ff4 & ", '" & NombreOperario(Me.p4) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.v4 <> 0 Then ' vacaciones
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (196, " & hv4 & ", #" & Me.fech & "#, " & Me.v4 & ", '" & NombreOperario(Me.p4) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.b4 <> 0 Then 'hay bono
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (109, 1, #" & Me.fech & "#, " & Me.b4 & ", '" & NombreOperario(Me.p4) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.i4 <> 0 Then 'hay inss
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (413, 1, #" & Me.fech & "#, " & Me.i4 & ", '" & NombreOperario(Me.p4) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Linea 5
If Me.s5 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (104, 1, #" & Me.fech & "#, " & Me.s5 & ", '" & NombreOperario(Me.p5) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.ff5 <> 0 Then ' feriado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (306, " & cf5 & ", #" & Me.fech & "#, " & Me.ff5 & ", '" & NombreOperario(Me.p5) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.v5 <> 0 Then ' vacaciones
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (196, " & hv5 & ", #" & Me.fech & "#, " & Me.v5 & ", '" & NombreOperario(Me.p5) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If
    
If Me.b5 <> 0 Then 'hay bono
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (109, 1, #" & Me.fech & "#, " & Me.b5 & ", '" & NombreOperario(Me.p5) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.i5 <> 0 Then 'hay inss
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (413, 1, #" & Me.fech & "#, " & Me.i5 & ", '" & NombreOperario(Me.p5) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Linea 6
If Me.s6 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (104, 1, #" & Me.fech & "#, " & Me.s6 & ", '" & NombreOperario(Me.p6) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.ff6 <> 0 Then ' feriado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (306, " & cf6 & ", #" & Me.fech & "#, " & Me.ff6 & ", '" & NombreOperario(Me.p6) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.v6 <> 0 Then ' vacaciones
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (196, " & hv6 & ", #" & Me.fech & "#, " & Me.v6 & ", '" & NombreOperario(Me.p6) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If
    
If Me.b6 <> 0 Then 'hay bono
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (109, 1, #" & Me.fech & "#, " & Me.b6 & ", '" & NombreOperario(Me.p6) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

If Me.i6 <> 0 Then 'hay inss
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (413, 1, #" & Me.fech & "#, " & Me.i6 & ", '" & NombreOperario(Me.p6) & "', 24, '" & "Pitaya " & codigoLocal() & "', 'OTROS', #" & Me.fech & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If
    
MsgBox "Salarios Agregados"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

