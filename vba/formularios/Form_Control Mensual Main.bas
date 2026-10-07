' ==========================================================
' Modulo  : Form_Control Mensual Main
' Tipo    : 100
' Lineas  : 53
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database
Private Sub actualizarlista()

Me.RecordSource = "SELECT nombreproductocoti([Cotizaciones]![CodCotizacion]) AS Nombre," & _
" ComprasMesMain([Cotizaciones]![CodCotizacion]," & Me.mesac & "," & Me.añoac & ") AS C, EgresosMesMain([Cotizaciones]![CodCotizacion]," & Me.mesac & "," & Me.añoac & ") AS I," & _
" StockFinalMesMain([Cotizaciones]![CodCotizacion]," & Me.mesac & "," & Me.añoac & ") AS SF, StockFinalMesMain([Cotizaciones]![CodCotizacion]," & Me.mesac & "," & Me.añoac & ") AS SI," & _
" DBIngredientes.Tipo, DBIngredientes.Control, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (DBIngredientes.Tipo = '" & Me.Tipo & "' And DBIngredientes.Control = True)" & _
" ORDER BY nombreproductocoti([Cotizaciones]![CodCotizacion])"

End Sub

Private Sub Comando100_Click()
DoCmd.OpenForm "Filtro Ingresos Semana"
[Forms]![Filtro Ingresos Semana]![bcodigo] = Me.CodCotizacion
End Sub

Private Sub Comando139_Click()
Me.mesac = IIf(Me.mesac = 1, 12, Me.mesac - 1)
Me.añoac = IIf(Me.mesac = 12, Me.añoac - 1, Me.añoac)
Me.mesanterior = IIf(Me.mesanterior = 1, 12, Me.mesanterior - 1)
Me.anoanterior = IIf(Me.mesanterior = 12, Me.anoanterior - 1, Me.anoanterior)

End Sub

Private Sub Comando142_Click()
Me.mesac = IIf(Me.mesac = 12, 1, Me.mesac + 1)
Me.añoac = IIf(Me.mesac = 1, Me.añoac + 1, Me.añoac)
Me.mesanterior = IIf(Me.mesanterior = 12, 1, Me.mesanterior + 1)
Me.anoanterior = IIf(Me.mesanterior = 1, anoanterior + 1, anoanterior)

End Sub

Private Sub Comando65_Click()
Call actualizarlista
End Sub

Private Sub Comando81_Click()
DoCmd.OpenForm "Filtro Compras Semana"
[Forms]![Filtro Compras Semana]![bcodigo] = Me.CodCotizacion
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 8000
Me.mesac = IIf(Month(Date) = 1, 12, Month(Date) - 1)
Me.añoac = IIf(Me.mesac = 1, Year(Date) - 1, Year(Date))
Me.mesanterior = IIf(Me.mesac = 1, 12, Me.mesac - 1)
Me.anoanterior = IIf(Me.mesac = 1, Me.añoac - 1, Me.añoac)
'Call actualizarlista
End Sub
