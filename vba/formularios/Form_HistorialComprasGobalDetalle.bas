' ==========================================================
' Modulo  : Form_HistorialComprasGobalDetalle
' Tipo    : 100
' Lineas  : 27
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

Me.InsideHeight = 8000
'Deifnir origen de base de datos base de datos mixed
Me.RecordSource = "SELECT Compras.CodCotizacion, Compras.CodProveedor, Compras.local," & _
" Compras.Pagado, Compras.Fecha, Compras.Destino," & _
" [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS Presentacion," & _
" Compras.CodIngresoAlmacen, Compras.Cantidad, Compras.CostoTotal, Compras.Observaciones," & _
" numerosemana([Compras]![Fecha]) AS semana, Year([Compras]![Fecha]) AS aano, Month([Compras]![Fecha]) AS ames," & _
" DBIngredientes.TIPO2, DBIngredientes.TIPO1, DBIngredientes.Tipo AS ctipo, Compras.Tipo, Proovedores.Nombre," & _
" [DBIngredientes]![NombreSinProcesar] & ' ' & [Cotizaciones]![Marca] & ' ' & [Cotizaciones]![Linea] & ' ' & [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS NombreProducto," & _
" Compras.CodProveedor, Compras.Destino, Compras.CodCotizacion, Compras.Pagado, Compras.Destino," & _
" Cotizaciones.CodIngrediente" & _
" FROM (DBIngredientes INNER JOIN (Cotizaciones INNER JOIN Compras" & _
" ON Cotizaciones.CodCotizacion = Compras.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente)" & _
" INNER JOIN Proovedores ON Compras.CodProveedor = Proovedores.CodProovedor" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" ORDER BY Compras.Fecha DESC , Compras.Destino DESC , Compras.Tipo DESC, Compras.CodProveedor DESC , " & _
" Compras.CodCotizacion, Compras.Pagado, Compras.Destino"
End Sub


