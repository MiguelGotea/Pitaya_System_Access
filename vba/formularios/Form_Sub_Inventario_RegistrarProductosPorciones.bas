' ==========================================================
' Modulo  : Form_Sub_Inventario_RegistrarProductosPorciones
' Tipo    : 100  |  Lineas: 50
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()

On Error GoTo AgainAgain
Dim tota As Double
Dim totaantiguo As Double

tota = InputBox("Ingrese la cantidad a registrar del siguiente producto: " & UCase(Me.Nombre), "Cantidad", 0)
totaantiguo = StockSinProcesar([codporcion], numerosemana([Forms]![Ingreso Inventario Pitaya]![fechaac]) + 1)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Cantidad, Fecha, CodOperario, lista)" & _
" values (" & Me.codporcion & ", " & tota - totaantiguo & ", #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ", 2)"
DoCmd.SetWarnings True

If ExisteMezcla(Me.codporcion) = 1 Then

    ' Obtener CodCotizacionPorcion relacionados
    Dim rsPorciones As DAO.Recordset
    Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & Me.codporcion)

    ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
    Do While Not rsPorciones.EOF
    
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion Temporal](CodCotizacion, Cantidad, Fecha, CodOperario, lista)" & _
            " values (" & rsPorciones!CodCotizacionPorcion & "," & tota - totaantiguo & ", #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ", 2)"
        DoCmd.SetWarnings False
        
        rsPorciones.MoveNext
    Loop

    rsPorciones.Close
End If
    
Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

