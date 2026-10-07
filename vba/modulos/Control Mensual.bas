' ==========================================================
' Modulo  : Control Mensual
' Tipo    : 1
' Lineas  : 95
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database

Function IngresosPitayaMes(codc As Integer, Mes As Integer, ano As Integer) As Long
'Ingresos totales de productos cotizacion  en el mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de cantidades ingresadas de cierto producto en mes especifico de ano especifico
miSQL = "SELECT IngresosPitaya.CodCotizacion, Sum(IngresosPitaya.Cantidad) AS SumaDeCantidad, Month([IngresosPitaya]![Fecha]) AS Mes, Year([IngresosPitaya]![Fecha]) AS Año FROM IngresosPitaya GROUP BY IngresosPitaya.CodCotizacion, Month([IngresosPitaya]![Fecha]), Year([IngresosPitaya]![Fecha]) HAVING (((IngresosPitaya.CodCotizacion)=" & codc & ") AND ((Month([IngresosPitaya]![Fecha]))=" & Mes & ") AND ((Year([IngresosPitaya]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresosPitayaMes = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
IngresosPitayaMes = 0

End Function

Function StockPitayaFinalMes(codc As Integer, Mes As Integer, ano As Integer) As Double
'Stock Final de productos cotizacion  en el mes especifico, stock final del mes presente tomado en fin demes = stock final con
' el que acaba el mes un poducto

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de cantidades stock final de cierto producto en mes especifico  de ano especifico
miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, Month([Inventario Cotizacion]![Fecha]) AS Mes, Year([Inventario Cotizacion]![Fecha]) AS Año, Sum([Inventario Cotizacion].Cantidad) AS SumaDeCantidad FROM [Inventario Cotizacion] GROUP BY [Inventario Cotizacion].CodCotizacion, Month([Inventario Cotizacion]![Fecha]), Year([Inventario Cotizacion]![Fecha]) HAVING ((([Inventario Cotizacion].CodCotizacion)=" & codc & ") AND ((Month([Inventario Cotizacion]![Fecha]))= " & Mes & ") AND ((Year([Inventario Cotizacion]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockPitayaFinalMes = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockPitayaFinalMes = 0

End Function

Function CostoServicioMensual(coti As Integer, Mes As Integer, ano As Integer, Cond As Integer) As Double
'Costo de servicio en todo el mes

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de costos de servicio en mes y ano especifico
miSQL = "SELECT Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Month([Compras]![Fecha]) AS Mes, Year([Compras]![Fecha]) AS Año, Compras.CodCotizacion, Max(Compras.Fecha) AS MáxDeFecha FROM DBIngredientes INNER JOIN (Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente GROUP BY Month([Compras]![Fecha]), Year([Compras]![Fecha]), Compras.CodCotizacion HAVING (((Month([Compras]![Fecha]))=" & Mes & ") AND ((Year([Compras]![Fecha]))=" & ano & ") AND ((Compras.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Cond = 1 Then
    CostoServicioMensual = rst("SumaDeCostoTotal")
    rst.Close
Else
    CostoServicioMensual = Day(rst("MáxDeFecha"))
    rst.Close
End If

Exit Function

Nulo:
CostoServicioMensual = 0

End Function

Function ConsumoFijoMensualIngrediente(ING As String, M As Integer, a As Integer) As Long
'COnsumo de un producto de costo fijo en un mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim ma, aa As Integer

ma = IIf(M = 1, 12, M - 1)
aa = IIf(M = 1, a - 1, a)

'Clacula de la suma de inventarios final inventario incial e ingresos
miSQL = "SELECT Sum(StockPitayaFinalMes([Cotizaciones]![CodCotizacion]," & ma & "," & aa & ")*DLookUp('[Conversion]','Cotizaciones','[CodCotizacion]=' & [Cotizaciones]![CodCotizacion])) AS INI, Sum(StockPitayaFinalMes([Cotizaciones]![CodCotizacion]," & M & "," & a & ")*DLookUp('[Conversion]','Cotizaciones','[CodCotizacion]=' & [Cotizaciones]![CodCotizacion])) AS FIN, Sum(IngresosPitayaMes([Cotizaciones]![CodCotizacion]," & M & "," & a & ")*DLookUp('[Conversion]','Cotizaciones','[CodCotizacion]=' & [Cotizaciones]![CodCotizacion])) AS ING, Cotizaciones.CodIngrediente FROM Cotizaciones GROUP BY Cotizaciones.CodIngrediente HAVING (((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoFijoMensualIngrediente = rst("INI") + rst("ING") - rst("FIN")
rst.Close

Exit Function

Nulo:
ConsumoFijoMensualIngrediente = 0

End Function



