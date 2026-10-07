' ==========================================================
' Modulo  : Form_Analisis de Productos
' Tipo    : 100
' Lineas  : 83
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database


Private Sub CodigoBusqueda_Exit(Cancel As Integer)
'MOSTRAR CODIGO DE INGREDIENTE
If Not IsNull(Me.CodigoBusqueda) Then
    Me.aingrediente.Value = DLookup("[CodIngrediente]", "Cotizaciones", "[CodCotizacion] = " & Me.CodigoBusqueda.Value)
End If


If Not Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Productos\" & Me.CodigoBusqueda & ".png") = "" Then
Me.imagenproducto.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Productos\" & Me.CodigoBusqueda & ".png"
Else
Me.imagenproducto.Picture = ""
End If
End Sub



Private Sub Comando366_Click()
DoCmd.OpenForm "Conv_Sem_Coti", acFormPivotChart, , "[CodCotizacion]=" & Me.CodigoBusqueda & " AND [semana]>" & Me.desde & " AND [semana]<" & numerosemana(Date)
End Sub



Private Sub Comando525_Click()
DoCmd.OpenForm "Consumo Ingredientes"
[Forms]![Consumo Ingredientes]![aingrediente] = Me.aingrediente
[Forms]![Consumo Ingredientes].Form.Requery
End Sub

Private Sub Comando530_Click()

DoCmd.OpenForm "Dashboard_Consumos_Semanales_Insumo", acFormPivotChart, , "[semana]>=" & Me.desde & " AND [semana]<=" & numerosemana(Date) & " AND [CodIngrediente]='" & Me.aingrediente & "'"
[Forms]![Dashboard_Consumos_Semanales_Insumo].Form.Requery

End Sub

Private Sub Comando533_Click()
DoCmd.OpenForm "Dashboard_Ingresos_Semanales_Insumo", acFormPivotChart, , "[semana]>=" & Me.desde & " AND [semana]<=" & numerosemana(Date) & " AND [CodIngrediente]='" & Me.aingrediente & "'"
[Forms]![Dashboard_Ingresos_Semanales_Insumo].Form.Requery
End Sub

Private Sub Comando539_Click()
DoCmd.OpenForm "Relacion de Productos Venta", , , "IngredienteContieneReceta([CodBatido],'" & Me.aingrediente & "')=1"
End Sub

Private Sub Comando542_Click()
DoCmd.OpenForm "Costo_Sem_Coti_Global", acFormPivotChart, , "[CodCotizacion]=" & Me.CodigoBusqueda & " AND [semana]>" & Me.desde & " AND [semana]<" & numerosemana(Date)
End Sub

Private Sub Comando545_Click()
DoCmd.OpenForm "Conv_Sem_Coti_Global", acFormPivotChart, , "[CodCotizacion]=" & Me.CodigoBusqueda & " AND [semana]>" & Me.desde & " AND [semana]<" & numerosemana(Date)

End Sub

Private Sub Comando548_Click()
DoCmd.OpenForm "HistorialComprasGobalDetalle", acNormal
[Forms]![HistorialComprasGobalDetalle].Form.Filter = "[CodIngrediente]='" & Me.aingrediente & "' AND [semana]>" & Me.desde & " AND [semana]<" & numerosemana(Date)
[Forms]![HistorialComprasGobalDetalle].Form.FilterOn = True
[Forms]![HistorialComprasGobalDetalle].Form.Requery
End Sub

Private Sub Comando599_Click()
DoCmd.OpenForm "Conv_Sem_Coti_xLocal", acFormPivotChart, , "[CodCotizacion]=" & Me.CodigoBusqueda & " AND [semana]>" & Me.desde & " AND [semana]<" & numerosemana(Date)

End Sub

Private Sub Comando727_Click()
DoCmd.OpenForm "ResumenPreIngresosPitaya", , , "[sema]>=" & Me.desde & " AND [CodCotizacion]=" & Me.CodigoBusqueda

End Sub

Private Sub Comando748_Click()
DoCmd.OpenForm "HistorialTransformacion", , , "[semana]>=" & Me.desde & " AND[CodIngrediente]='" & Me.aingrediente & "'"
[Forms]![HistorialTransformacion].Form.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub
