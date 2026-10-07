' ==========================================================
' Modulo  : Form_Control Diario Compra Ingreso Global
' Tipo    : 100  |  Lineas: 149
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando117_Click() 'Marca
If Me.amarca = "" Then
    MsgBox "Ingrese Marca"
    Exit Sub
End If

Me.variacion.width = 1000
Me.Local.width = 0
Me.Cancelado.width = 0

Me.Fecha.Visible = True
Me.I.Visible = True
Me.C.Visible = True
Me.Local.Visible = True
Me.CodCotizacion.Visible = True
Me.variacion.Visible = True
Me.aname.Visible = True
Me.asemana.Visible = True
Me.Cancelado.Visible = True

Me.aproducto = ""

Me.RecordSource = "SELECT [U_IngresoxCompra]![Fecha] Between #" & Me.desde & "# And #" & Me.hasta & "# AS Periodo, DBIngredientes.TIPO1, DBIngredientes.Control, U_IngresoxCompra.CodCotizacion, Sum(IIf([U_IngresoxCompra]![tipo]='C',[U_IngresoxCompra]![Cantidad],0)) AS C, Sum(IIf([U_IngresoxCompra]![tipo]='I',[U_IngresoxCompra]![Cantidad],0)) AS I, 0 AS [local], Min(U_IngresoxCompra.Fecha) AS Fecha, Cotizaciones.Marca" & _
" FROM DBIngredientes INNER JOIN (U_IngresoxCompra INNER JOIN Cotizaciones ON U_IngresoxCompra.CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY [U_IngresoxCompra]![Fecha] Between #" & Me.desde & "# And #" & Me.hasta & "#, DBIngredientes.TIPO1, DBIngredientes.Control, U_IngresoxCompra.CodCotizacion, 0, Cotizaciones.Marca" & _
" HAVING ((([U_IngresoxCompra]![Fecha] Between #" & Me.desde & "# And #" & Me.hasta & "#)<>0) AND ((DBIngredientes.TIPO1)='" & Me.Tipo & "') AND ((DBIngredientes.Control) = True) AND ((Cotizaciones.Marca) = '" & Me.amarca & "'))" & _
" ORDER BY U_IngresoxCompra.CodCotizacion"

End Sub

Private Sub Comando8_Click() 'dia por dia producto
If Me.aproducto = "" Then
    MsgBox "Ingrese Producto"
    Exit Sub
End If

Me.variacion.width = 0
Me.Local.width = 1000
Me.Cancelado.width = 1500

Me.Fecha.Visible = True
Me.I.Visible = True
Me.C.Visible = True
Me.Local.Visible = True
Me.CodCotizacion.Visible = True
Me.variacion.Visible = True
Me.aname.Visible = True
Me.asemana.Visible = True
Me.Cancelado.Visible = True

Me.RecordSource = "SELECT U_IngresoxCompra.CodCotizacion, U_IngresoxCompra.Fecha, U_IngresoxCompra.local, IIf([U_IngresoxCompra]![tipo]='C',[U_IngresoxCompra]![Cantidad],0) AS C, IIf([U_IngresoxCompra]![tipo]='I',[U_IngresoxCompra]![Cantidad],0) AS I, U_IngresoxCompra.cancelado" & _
" FROM U_IngresoxCompra IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((U_IngresoxCompra.CodCotizacion)=" & Me.aproducto & ") AND ((U_IngresoxCompra.Fecha) Between #" & Me.desde & "# And #" & Me.hasta & "#))" & _
" ORDER BY U_IngresoxCompra.Fecha ASC, IIf([U_IngresoxCompra]![tipo]='C',[U_IngresoxCompra]![Cantidad],0) DESC"
End Sub

Private Sub Comando103_Click() 'Por Tipo
If Me.Tipo = "" Then
    MsgBox "Ingrese Tipo"
    Exit Sub
End If

Me.variacion.width = 1000
Me.Local.width = 0
Me.Cancelado.width = 0

Me.Fecha.Visible = True
Me.I.Visible = True
Me.C.Visible = True
Me.Local.Visible = True
Me.CodCotizacion.Visible = True
Me.variacion.Visible = True
Me.aname.Visible = True
Me.asemana.Visible = True
Me.Cancelado.Visible = True

Me.aproducto = ""
Me.amarca = ""

Me.RecordSource = "SELECT [U_IngresoxCompra]![Fecha] Between #" & Me.desde & "# And #" & Me.hasta & "# AS Periodo, DBIngredientes.TIPO1, DBIngredientes.Control, U_IngresoxCompra.CodCotizacion, Sum(IIf([U_IngresoxCompra]![tipo]='C',[U_IngresoxCompra]![Cantidad],0)) AS C, Sum(IIf([U_IngresoxCompra]![tipo]='I',[U_IngresoxCompra]![Cantidad],0)) AS I, Min(U_IngresoxCompra.Fecha) AS Fecha, nombreproductocoti([U_IngresoxCompra].[CodCotizacion]) As Nombre, Cotizaciones.Global" & _
" FROM DBIngredientes INNER JOIN (U_IngresoxCompra INNER JOIN Cotizaciones ON U_IngresoxCompra.CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY [U_IngresoxCompra]![Fecha] Between #" & Me.desde & "# And #" & Me.hasta & "#, DBIngredientes.TIPO1, DBIngredientes.Control, U_IngresoxCompra.CodCotizacion, nombreproductocoti([U_IngresoxCompra].[CodCotizacion]), Cotizaciones.Global" & _
" HAVING ((([U_IngresoxCompra]![Fecha] Between #" & Me.desde & "# And #" & Me.hasta & "#)<>0) AND ((DBIngredientes.TIPO1)='" & Me.Tipo & "') AND ((DBIngredientes.Control) = True))" & _
" ORDER BY nombreproductocoti([U_IngresoxCompra].[CodCotizacion])"

If Me.aglobal = True Then
    Me.Filter = "[Global]<>0"
    Me.FilterOn = True
Else
    Me.Filter = "[Global]=0"
    Me.FilterOn = True
End If
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 9000
Me.desde = Date
Me.hasta = Date
Me.aglobal = True

Me.Fecha.Visible = False
Me.C.Visible = False
Me.I.Visible = False
Me.Local.Visible = False
Me.CodCotizacion.Visible = False
Me.variacion.Visible = False
Me.aname.Visible = False
Me.asemana.Visible = False
Me.Cancelado.Visible = False

End Sub

Private Sub tipo_Change()
Me.aproducto = ""
Me.amarca = ""
Me.aproducto.Requery
Me.amarca.Requery
End Sub

Private Sub amarca_Change()
Me.aproducto = ""
End Sub

Private Sub aproducto_Change()
Me.amarca = ""
End Sub

Private Sub Comando33_Click()
DoCmd.OpenForm "Filtro Compras Semana"
End Sub

Private Sub Comando39_Click()
DoCmd.OpenForm "Filtro Ingresos Semana"
End Sub

Private Sub variacion_Click()
Me.aproducto = Me.CodCotizacion
Me.amarca = MarcaProducto(Me.CodCotizacion)
End Sub

Private Sub CodCotizacion_Click()
Me.aproducto = Me.CodCotizacion
Me.amarca = MarcaProducto(Me.CodCotizacion)
End Sub
