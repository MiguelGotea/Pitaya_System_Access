' ==========================================================
' Modulo  : Form_AutoRegistroPagoServicios
' Tipo    : 100  |  Lineas: 114
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Private Sub Comando39_Click()
'Energia
If Me.b1 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (100, 1, #" & Me.s1 & "#, " & Me.b1 & ", ' ', 27, '" & "Pitaya " & codigoLocal() & "', 'TIPO 5', #" & Me.s1 & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Agua
If Me.b2 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (99, 1, #" & Me.s2 & "#, " & Me.b1 & ", ' ', 29, '" & "Pitaya " & codigoLocal() & "', 'TIPO 5', #" & Me.s2 & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Internet
If Me.b3 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (101, 1, #" & Me.s3 & "#, " & Me.b3 & ", ' ', 28, '" & "Pitaya " & codigoLocal() & "', 'TIPO 5', #" & Me.s3 & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If


'Tren Aseo
If Me.b4 <> 0 Then ' hay valor apra registrar
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (174, 1, #" & Me.s4 & "#, " & Me.b4 & ", ' ', 40, '" & "Pitaya " & codigoLocal() & "', 'TIPO 5', #" & Me.s4 & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If

'Renta
If Me.b5 <> 0 Then ' hay valor apra registrar
    Dim loca, proove As Integer
    loca = codigoLocal()
    If loca = 2 Then proove = 23 Else proove = 60
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Compras(CodCotizacion, Cantidad, Fecha, CostoTotal, Observaciones, CodProveedor, Destino, Tipo, Pagado, Hora) values" & _
    " (105, 1, #" & Me.s5 & "#, " & Me.b5 & ", ' ', " & proove & ", '" & "Pitaya " & codigoLocal() & "', 'TIPO 5', #" & Me.s5 & "#, #" & Time() & "#)"
    DoCmd.SetWarnings True
End If
    
MsgBox "Facturas agregadas"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

Private Sub aano_Exit(Cancel As Integer)
Call actualizarfechas
End Sub

Private Sub ames_Exit(Cancel As Integer)
Call actualizarfechas
End Sub

Private Sub actualizarfechas()
Dim cmes As Integer
Dim Quincena, finmes As Date

If IsNull(Me.ames) Then
    MsgBox "agregar mes"
    Exit Sub
End If
If IsNull(Me.aano) Then
    MsgBox "agregar año"
    Exit Sub
End If

Select Case Me.ames
    Case "ENERO"
        cmes = 1
    Case "FEBRERO"
        cmes = 2
    Case "MARZO"
        cmes = 3
    Case "ABRIL"
        cmes = 4
    Case "MAYO"
        cmes = 5
    Case "JUNIO"
        cmes = 6
    Case "JULIO"
        cmes = 7
    Case "AGOSTO"
        cmes = 8
    Case "SETIEMBRE"
        cmes = 9
    Case "OCTUBRE"
        cmes = 10
    Case "NOVIEMBRE"
        cmes = 11
    Case "DICIEMBRE"
        cmes = 12
End Select

Quincena = DateSerial(Me.aano, cmes, 15)
finmes = DateSerial(Me.aano, cmes + 1, 0)

Me.s1 = Quincena
Me.s2 = Quincena
Me.s3 = finmes
Me.s4 = finmes
Me.s5 = Quincena
End Sub

