' ==========================================================
' Modulo  : Form_InventarioDisponiblePorcionarAntiguo
' Tipo    : 100  |  Lineas: 71
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database

Private Sub Comando16_Click()
DoCmd.OpenForm "Porcionamiento"

[Forms]![Porcionamiento]![cing] = Me.CodIngrediente
[Forms]![Porcionamiento]![cproce] = 0 'COdigo de Procesamiento
[Forms]![Porcionamiento]![cprocedencia] = Me.CodCotizacion
[Forms]![Porcionamiento]![nprocedencia] = nombreproductocoti(Me.CodCotizacion)
[Forms]![Porcionamiento]![aingrediente] = DLookup("[Nombre]", "DBIngredientes", "[CodIngrediente]='" & Me.CodIngrediente & "'")
[Forms]![Porcionamiento]![convproce] = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion]=" & Me.CodCotizacion & "")
[Forms]![Porcionamiento]![uniing] = DLookup("[Unidad]", "DBIngredientes", "[CodIngrediente]='" & Me.CodIngrediente & "'")
[Forms]![Porcionamiento]![unicoti] = DLookup("[Unidad]", "Cotizaciones", "[CodCotizacion]=" & Me.CodCotizacion & "")

If Me.asemana = numerosemana(Date) Then
    [Forms]![Porcionamiento]![cfecha] = Date
    [Forms]![Porcionamiento]![csemana] = numerosemana(Date)
    [Forms]![Porcionamiento].Form.Filter = "[Procedencia]=" & Me.CodCotizacion & " AND [semana]=" & numerosemana(Date)
Else
    [Forms]![Porcionamiento]![csemana] = Me.asemana
    [Forms]![Porcionamiento].Form.Filter = "[Procedencia]=" & Me.CodCotizacion & " AND [semana]=" & Me.asemana
End If

[Forms]![Porcionamiento].Form.FilterOn = True
[Forms]![Porcionamiento].Requery
End Sub

Private Sub Comando21_Click()
Me.Requery
End Sub

Private Sub Comando25_Click()
Me.asemana.Value = Me.asemana.Value + 1
Me.Requery
End Sub

Private Sub Comando267_Click()
If Me.Comando267.Caption = "OCULTAR PORCIONES GLOBAL" Then
    Me.Filter = "([Marca]<>'Almacen Global' or isnull([Marca])<>0) and ([Ingresos] <> 0 or [SI] <> 0 or [Porcionado] <> 0) "
    Me.FilterOn = True
    Me.Requery
    Me.Comando267.Caption = "MOSTRAR TODO"
Else
    Me.Filter = ""
    Me.FilterOn = True
    Me.Requery
    Me.Comando267.Caption = "OCULTAR PORCIONES GLOBAL"
End If
End Sub

Private Sub Comando28_Click()
Me.asemana.Value = Me.asemana.Value - 1
Me.Requery
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaactual = numerosemana(Date)
Me.asemana = numerosemana(Date)
Me.Requery

Me.Filter = "([Marca]<>'Almacen Global' or isnull([Marca])<>0) and ([Ingresos] <> 0 or [SI] <> 0 or [Porcionado] <> 0)"
Me.FilterOn = True

Me.Requery
End Sub
