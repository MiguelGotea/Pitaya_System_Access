' ==========================================================
' Modulo  : Form_SolicitudDespachoSucursales
' Tipo    : 100  |  Lineas: 189
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database

Private Sub generarpreingresosegunpedido(preingresoprimer As Long, preingresosegundo As Long)
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim rsPorciones As DAO.Recordset
Dim canti As Integer
Dim ind As Integer
Dim prei As Long
Dim conver As Integer

miSQL = "SELECT [Inventario Cotizacion].CodICotizacion, [Inventario Cotizacion].CodCotizacion," & _
" [Inventario Cotizacion].primerenvio, [Inventario Cotizacion].segundoenvio," & _
" numerosemana([Inventario Cotizacion]![Fecha])+1 AS semanasolicitud" & _
" FROM [Inventario Cotizacion]" & _
" WHERE (((numerosemana([Inventario Cotizacion]![Fecha]) + 1) = " & Me.asemana & "))" & _
" ORDER BY [Inventario Cotizacion].CodICotizacion"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

'conver = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & rst("CodCotizacion"))
conver = factorinventariounidaddespacho(rst("CodCotizacion"))

For ind = 1 To canti

    If ind >= Me.icoti Then
        'MsgBox nombreproductocoti(rst("CodCotizacion"))
        conver = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & rst("CodCotizacion"))

        If rst("primerenvio") <> 0 Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & rst("CodCotizacion") & ", " & rst("primerenvio") * conver & ", " & preingresoprimer & ")"
            DoCmd.SetWarnings True
            
            If ExisteMezcla(rst("CodCotizacion")) = 1 Then
                ' Agregar la mezcla original
                ' Obtener CodCotizacionPorcion relacionados

                Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & rst("CodCotizacion"))
        
                ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
                Do While Not rsPorciones.EOF
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
                        " values (" & rsPorciones!CodCotizacionPorcion & ", " & rst("primerenvio") * conver & ", " & preingresoprimer & ")"
                    rsPorciones.MoveNext
                    DoCmd.SetWarnings True
                Loop
        
                rsPorciones.Close
            End If

        End If
        
        If rst("segundoenvio") <> 0 Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & rst("CodCotizacion") & ", " & rst("segundoenvio") * conver & ", " & preingresosegundo & ")"
            DoCmd.SetWarnings True
            
            If ExisteMezcla(rst("CodCotizacion")) = 1 Then
                ' Agregar la mezcla original
                ' Obtener CodCotizacionPorcion relacionados
                
                Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & rst("CodCotizacion"))
        
                ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
                Do While Not rsPorciones.EOF
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
                        " values (" & rsPorciones!CodCotizacionPorcion & ", " & rst("segundoenvio") * conver & ", " & preingresosegundo & ")"
                    rsPorciones.MoveNext
                    DoCmd.SetWarnings True
                Loop
        
                rsPorciones.Close
            End If

        End If
    End If
    rst.MoveNext
Next ind


rst.Close
Me.icoti = 1
Exit Sub

AgainAgain:
Me.icoti = ind

'MsgBox nombreproductocoti(rst("CodCotizacion"))
rst.Close
Call generarpreingresosegunpedido(preingresoprimer, preingresosegundo)


End Sub





Private Sub Comando155_Click()
On Error Resume Next
Me.RecordSource = ""
Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "Inventario Cotizacion", "Inventario Cotizacion", 3)
Me.RecordSource = "SELECT [Inventario Cotizacion].CodICotizacion, [Inventario Cotizacion].CodCotizacion," & _
" [Inventario Cotizacion].primerenvio, [Inventario Cotizacion].segundoenvio, numerosemana([Inventario Cotizacion]![Fecha])+1 AS semanasolicitud," & _
" factorinventariounidaddespacho([Inventario Cotizacion]![CodCotizacion]) AS conver" & _
" FROM [Inventario Cotizacion]" & _
" WHERE (((numerosemana([Inventario Cotizacion]![Fecha])+1)=[Formularios]![SolicitudDespachoSucursales]![asemana]))" & _
" ORDER BY [Inventario Cotizacion].CodICotizacion"
Me.Requery
End Sub



Private Sub Comando939_Click()
Call Comando155_Click

If MsgBox("Desea generar preingresos segun los datos enviados por las sucursales?", vbYesNo, "CONFIRMACION") = vbYes Then
    'se sube el hroario
Else
    Exit Sub
End If

If IsNull(Me.Fecha1) Then
    MsgBox "Ingresar Fechas de despacho"
    Exit Sub
End If

If IsNull(Me.Fecha2) Then
    MsgBox "Ingresar Fechas de despacho"
    Exit Sub
End If

Me.Requery

Dim primeren As Date
Dim segundoen As Date
Dim primerpre As Long
Dim segundopre As Long

primeren = Me.Fecha1
segundoen = Me.Fecha2

primerpre = crearpreingreso(primeren, Me.asucursal)
segundopre = crearpreingreso(segundoen, Me.asucursal)

Call generarpreingresosegunpedido(primerpre, segundopre)

MsgBox "Listas generadas"

End Sub





Private Sub Form_Close()
Me.RecordSource = ""
Call importartablaespecifica("Almacen", "Inventario Cotizacion", "Inventario Cotizacion", 3)
Me.RecordSource = "SELECT [Inventario Cotizacion].CodICotizacion, [Inventario Cotizacion].CodCotizacion," & _
" [Inventario Cotizacion].primerenvio, [Inventario Cotizacion].segundoenvio, numerosemana([Inventario Cotizacion]![Fecha])+1 AS semanasolicitud," & _
" factorinventariounidaddespacho([Inventario Cotizacion]![CodCotizacion]) AS conver" & _
" FROM [Inventario Cotizacion]" & _
" WHERE (((numerosemana([Inventario Cotizacion]![Fecha])+1)=[Formularios]![SolicitudDespachoSucursales]![asemana]))" & _
" ORDER BY [Inventario Cotizacion].CodICotizacion"
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.RecordSource = ""
Call importartablaespecifica("Almacen", "Inventario Cotizacion", "Inventario Cotizacion", 3)
Me.RecordSource = "SELECT [Inventario Cotizacion].CodICotizacion, [Inventario Cotizacion].CodCotizacion," & _
" [Inventario Cotizacion].primerenvio, [Inventario Cotizacion].segundoenvio, numerosemana([Inventario Cotizacion]![Fecha])+1 AS semanasolicitud," & _
" factorinventariounidaddespacho([Inventario Cotizacion]![CodCotizacion]) AS conver" & _
" FROM [Inventario Cotizacion]" & _
" WHERE (((numerosemana([Inventario Cotizacion]![Fecha])+1)=[Formularios]![SolicitudDespachoSucursales]![asemana]))" & _
" ORDER BY [Inventario Cotizacion].CodICotizacion"
Me.Requery
End Sub


