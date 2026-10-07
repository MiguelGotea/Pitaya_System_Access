' ==========================================================
' Modulo  : Form_Procesamiento de Insumos
' Tipo    : 100
' Lineas  : 70
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
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
'Crear SubPorcionamiento si no hay
Dim codsubpor As Long
Dim proces As Long
Dim inger As String
Dim cotix As Integer
Dim cuanti As Double

cotix = Me.CodCotizacion
inger = Me.CodIngrediente
proces = Me.CodProcesamiento
codsubpor = BuscarCodSubporcionamiento(proces)
cuanti = Me.Cantidad

If codsubpor = 0 Then
    MsgBox "Se ha creado un Porcionamiento de este insumo"
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO SubPorcionamiento(Procedencia, CodProcesamiento, Cantidad, Fecha) values" & _
    " (" & Me.CodCotizacion & ", " & proces & ", " & cuanti & ", #" & Me.Fecha & "#)"
    DoCmd.SetWarnings True
    Me.Requery
    codsubpor = BuscarCodSubporcionamiento(proces)
End If


DoCmd.OpenForm "Porcionamiento"
[Forms]![Porcionamiento]![cing] = inger
[Forms]![Porcionamiento]![cfecha] = Me.Fecha
[Forms]![Porcionamiento]![csemana] = numerosemana(Me.Fecha)
[Forms]![Porcionamiento]![cproce] = proces
[Forms]![Porcionamiento]![cprocedencia] = cotix
[Forms]![Porcionamiento]![nprocedencia] = nombreproductocoti(cotix)
[Forms]![Porcionamiento]![asubporcionamiento] = codsubpor
[Forms]![Porcionamiento].Requery
End Sub


Private Sub Comando268_Click()
DoCmd.OpenForm "IngresoAutomaticoProductosFiltrado", acNormal, , "[Conversion]=0"

[Forms]![IngresoAutomaticoProductosFiltrado]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductosFiltrado]![adestino] = "[Procesamiento]"
[Forms]![IngresoAutomaticoProductosFiltrado]![adesdeform] = "Procesamiento de Insumos"
[Forms]![IngresoAutomaticoProductosFiltrado]![aproducto].Visible = False
[Forms]![IngresoAutomaticoProductosFiltrado]![Etiqueta153].Visible = False
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub
