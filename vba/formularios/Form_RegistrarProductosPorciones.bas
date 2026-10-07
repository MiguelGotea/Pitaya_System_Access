' ==========================================================
' Modulo  : Form_RegistrarProductosPorciones
' Tipo    : 100  |  Lineas: 78
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()
On Error GoTo AgainAgain
Dim tota As Double
Dim rsPorciones As DAO.Recordset

tota = InputBox("Ingrese la cantidad de productos a registrar", "Cantidad", 0)

DoCmd.SetWarnings False

If Me.adesdetabla = "[SubPreIngresosPitaya]" Or Me.adesdetabla = "[CambiosPreIngresosPitaya]" Then
    ' Verificar si existe relación con CodCotizacionPorcion

    DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
        " values (" & Me.codporcion & ", " & tota & ", " & Me.apreingre & ")"
        
    If ExisteMezcla(Me.codporcion) = 1 Then
        ' Agregar la mezcla original
        ' Obtener CodCotizacionPorcion relacionados
        
        Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & Me.codporcion)

        ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
        Do While Not rsPorciones.EOF
            DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
                " values (" & rsPorciones!CodCotizacionPorcion & ", " & tota & ", " & Me.apreingre & ")"
            rsPorciones.MoveNext
        Loop

        rsPorciones.Close
    End If
    
Else
    DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, Fecha)" & _
    " values (" & Me.codporcion & ", " & tota & ", #" & Me.fechaprocedencia & "#)"

    If ExisteMezcla(Me.codporcion) = 1 Then
        ' Agregar la mezcla original
        ' Obtener CodCotizacionPorcion relacionados

        Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & Me.codporcion)

        ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
        Do While Not rsPorciones.EOF
            DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, Fecha)" & _
                " values (" & rsPorciones!CodCotizacionPorcion & ", " & tota & ", #" & Me.fechaprocedencia & "#)"
            rsPorciones.MoveNext
        Loop

        rsPorciones.Close
    End If
    
End If

DoCmd.SetWarnings True

Me.icoti = 1
Me.guardado.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub

Private Sub Form_Close()
If Me.adesdeform = "[Ingreso Inventario Pitaya]" Then
    Forms(Me.adesdeform).Form.Subformulario_Inventario_Cotizacion.Requery
Else
    Forms(Me.adesdeform).Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
