' ==========================================================
' Modulo  : Form_Porcionamiento de Insumos
' Tipo    : 100  |  Lineas: 69
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database



Private Sub Comando150_Click()
Me.fechaac.Value = Me.fechaac.Value - 1
Me.Requery
End Sub

Private Sub Comando151_Click()
Me.fechaac.Value = Me.fechaac.Value + 1
Me.Requery
End Sub

Private Sub Comando159_Click()
DoCmd.OpenForm "Porcionamiento"
[Forms]![Porcionamiento]![cing] = Me.CodIngrediente
[Forms]![Porcionamiento]![pfecha] = Me.Fecha
[Forms]![Porcionamiento]![cantidadprocesar] = Me.Cantidad

[Forms]![Porcionamiento]![csemana] = numerosemana(Me.Fecha)
[Forms]![Porcionamiento]![cproce] = Me.CodProcesamiento
[Forms]![Porcionamiento]![cprocedencia] = Me.Procedencia
[Forms]![Porcionamiento]![asubporcionamiento] = Me.CodSubPorcionamiento
[Forms]![Porcionamiento]![nprocedencia] = nombreproductocoti(Me.Procedencia)


[Forms]![Porcionamiento]![CodCotizacion].RowSource = "SELECT nombreproductocotiprocesado([Cotizaciones]![CodCotizacion]) AS Nombre," & _
" Cotizaciones.CodCotizacion, Cotizaciones.Descontinuado, Cotizaciones.CodIngrediente, Cotizaciones.Subproducto, Cotizaciones.Prioridad, [Cotizaciones]![Marca] & ' ' As Marca2" & _
" FROM Cotizaciones" & _
" WHERE (((Cotizaciones.CodIngrediente)='" & [Forms]![Porcionamiento]![cing] & "')" & _
" AND (([Cotizaciones]![Marca] & ' ') <> 'Almacen Global ')" & IIf(codigoLocal() <> 0, " AND ((Cotizaciones.Prioridad)<>0)", "") & " AND ((Cotizaciones.Descontinuado)=0) AND ((Cotizaciones.CodCotizacion)<>" & Me.Procedencia & "))" & _
" ORDER BY nombreproductocotiprocesado([Cotizaciones]![CodCotizacion])"


[Forms]![Porcionamiento].Requery
End Sub


Private Sub Comando268_Click()
DoCmd.OpenForm "IngresoAutomaticoProductosFiltrado"
[Forms]![IngresoAutomaticoProductosFiltrado]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductosFiltrado]![adestino] = "[SubPorcionamiento]"
[Forms]![IngresoAutomaticoProductosFiltrado]![adesdeform] = "[Porcionamiento de Insumos]"
[Forms]![IngresoAutomaticoProductosFiltrado]![modulo] = "produccion"
End Sub

Private Sub Comando322_Click()
DoCmd.OpenForm "IngresoAutomaticoProductosFiltrado", , , "[Conversion]=0"
[Forms]![IngresoAutomaticoProductosFiltrado]![Etiqueta153].Visible = False
[Forms]![IngresoAutomaticoProductosFiltrado]![aproducto].Visible = False
[Forms]![IngresoAutomaticoProductosFiltrado]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductosFiltrado]![adestino] = "[SubPorcionamiento]"
[Forms]![IngresoAutomaticoProductosFiltrado]![adesdeform] = "[Porcionamiento de Insumos]"
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
If codigoLocal() = 0 Then
    Me.Comando322.Visible = True
End If

Me.Requery
End Sub
