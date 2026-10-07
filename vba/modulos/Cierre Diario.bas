' ==========================================================
' Modulo  : Cierre Diario
' Tipo    : 1  |  Lineas: 445
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database

Function SelladorInicial(Fecha As Date) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Contador Inicial de selladora
miSQL = "SELECT EstadoInicial.Fecha, EstadoInicial.Selladora FROM EstadoInicial WHERE (((EstadoInicial.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SelladorInicial = rst("Selladora")
rst.Close

Exit Function

Nulo:
MsgBox "No Hay Datos Iniciales ALmacenados"
SelladorInicial = 0

End Function

Function UltimoCierre() As Date

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Ultimo registro de cierre guardado, fecha y hora
miSQL = "SELECT CierreDiario.HoraFinal, CierreDiario.Fecha FROM CierreDiario GROUP BY CierreDiario.HoraFinal, CierreDiario.Fecha HAVING (((CierreDiario.Fecha) = #" & Date & "#)) ORDER BY CierreDiario.HoraFinal DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
UltimoCierre = rst("HoraFinal")
rst.Close

Exit Function

Nulo:
UltimoCierre = #5:00:00 AM#

End Function
Function CierreFinal(fech As Date) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar el codigo del ultimo cierre del dia especifico
miSQL = "SELECT CierreDiario.CodigoCierre, CierreDiario.HoraFinal, CierreDiario.Fecha FROM CierreDiario WHERE (((CierreDiario.Fecha) = #" & fech & "#)) ORDER BY CierreDiario.HoraFinal DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CierreFinal = rst("CodigoCierre")
rst.Close

Exit Function

Nulo:
CierreFinal = 0

End Function
Function cajafinal(Fecha As Date) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma del monto del  periodo + caja inicial para dar caja final del periodo
miSQL = "SELECT EstadoInicial.Fecha, AcumuladoDia([EstadoInicial]![Fecha])+[EstadoInicial]![Dinero] AS Monto FROM EstadoInicial GROUP BY EstadoInicial.Fecha, AcumuladoDia([EstadoInicial]![Fecha])+[EstadoInicial]![Dinero] HAVING (((EstadoInicial.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cajafinal = rst("Monto")
rst.Close

Exit Function

Nulo:
cajafinal = 0

End Function
Function MontoDelPeriodo(Fecha As Date) As Long
' SUma de monto del ultimo periodo de cierre que sirve para calcular ventas de periodo
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma del monto del ultimo periodo
miSQL = "SELECT NotaDePedido.Fecha, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto, [Hora]<= #" & Time() & "# And [Hora]> #" & UltimoCierre() & "# AS Condicional FROM NotaDePedido GROUP BY NotaDePedido.Fecha, [Hora]<=#" & Time() & "# And [Hora]>#" & UltimoCierre() & "# HAVING (((NotaDePedido.Fecha)=#" & Fecha & "#) AND (([Hora]<=#" & Time() & "# And [Hora]>#" & UltimoCierre() & "#)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MontoDelPeriodo = rst("Monto")
rst.Close

Exit Function

Nulo:
MontoDelPeriodo = 0

End Function
Function ContadorFinal(Fecha As Date) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'COntador Final de maquina selladora
miSQL = "SELECT NotaDePedido.Fecha, Sum(ConsumoEmpaque([SubPedido]![CodSubPedido],'Kid',0)+ConsumoEmpaque([SubPedido]![CodSubPedido],'Gigantona',0)+ConsumoEmpaque([SubPedido]![CodSubPedido],'Mediano',0)+ConsumoEmpaque([SubPedido]![CodSubPedido],'Bowl',0)) AS Consumo FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido GROUP BY NotaDePedido.Fecha HAVING (((NotaDePedido.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ContadorFinal = rst("Consumo") + SelladorInicial(Fecha)
rst.Close

Exit Function

Nulo:
ContadorFinal = SelladorInicial(Fecha)

End Function

Function MontoCierre(Fecha As Date, deno As String) As Double
'Dolares (D)
'Cordobas (C)
'Total (T)
'cierre de una fecha especifica de un local especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcula de manera descendente los cierres de fecha especifica y se extrae los dol o cordobas del cierre
miSQL = "SELECT CierreDiario.HoraFinal, CierreDiario.Fecha, CierreDiario.MFCor, CierreDiario.MFDol FROM CierreDiario WHERE (((CierreDiario.Fecha) = #" & Fecha & "#)) ORDER BY CierreDiario.HoraFinal DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Select Case deno
    Case "C" 'Inicio de semanas 1 el 1/1/2016
        MontoCierre = rst("MFCor")
    Case "D"
        MontoCierre = rst("MFDol")
    Case "T"
        MontoCierre = rst("MFDol") * tipocambio(Fecha) + rst("MFCor")
    Case Else
        MsgBox "No existe denominacion"
        MontoCierre = 0
End Select
    
rst.Close

Exit Function

Nulo:
MontoCierre = 0

End Function

Function MontoDeposito(Fecha As Date, deno As String) As Double
'Monto Depositado en una fecha especifica con una denominacion especifica
'deno= cordobas o dolares

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcula Suma de monto de la misma denominacion de una fecha especifica
miSQL = "SELECT Depositos.Fecha, Depositos.Denominacion, Depositos.DuranteTurno, Sum(Depositos.Monto) AS SumaDeMonto" & _
" FROM Depositos GROUP BY Depositos.Fecha, Depositos.Denominacion, Depositos.DuranteTurno" & _
" HAVING (((Depositos.Fecha)=#" & Fecha & "#) AND ((Depositos.Denominacion)='" & deno & "') AND ((Depositos.DuranteTurno)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MontoDeposito = rst("SumaDeMonto")
rst.Close

Exit Function

Nulo:
MontoDeposito = 0

End Function

Function cajainicial(Fecha As Date) As Double
'Monto Caja Inicial de fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcula caja inicial de fecha especifica
miSQL = "SELECT EstadoInicial.Fecha, EstadoInicial.Dinero FROM EstadoInicial WHERE (((EstadoInicial.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cajainicial = rst("Dinero")
rst.Close

Exit Function

Nulo:
cajainicial = 0

End Function

Function pagospordia(Fecha As Date) As Double
'Monto Pagado en total en cada dia acorde a la fecha de pago de cada factura

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcula Suma de monto de las compras pagadas en una fecha especifica
miSQL = "SELECT Compras.Fecha, Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Compras.Tipo," & _
" IsNull([Compras]![Fecha]) AS nulo FROM Compras GROUP BY Compras.Fecha, Compras.Tipo," & _
" IsNull([Compras]![Fecha]) HAVING (((Compras.Fecha)=#" & Fecha & "#) AND ((Compras.Tipo)='CAJA')" & _
" AND ((IsNull([Compras]![Fecha]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pagospordia = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
pagospordia = 0

End Function


Function tipocambio(bfech As Date) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Tipo de Cambio de Local Pitaya
miSQL = "SELECT TipoCambio.Fecha, TipoCambio.Monto FROM tipocambio WHERE (((TipoCambio.Fecha)=#" & bfech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
tipocambio = rst("Monto")
rst.Close

Exit Function

Nulo:
tipocambio = 0

End Function


Function AcumuladoCierresDiarios(desde As Date, hasta As Date) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim fech As Date

'Tipo de Cambio de Local Pitaya
miSQL = "SELECT FechaSistema.Dates FROM FechaSistema" & _
" WHERE (((FechaSistema.Dates)<=#" & hasta & "# And (FechaSistema.Dates)>=#" & desde & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst
AcumuladoCierresDiarios = 0

For I = 1 To Cant
    fech = rst("Dates")
    AcumuladoCierresDiarios = AcumuladoCierresDiarios + MontoCierre(fech, "T") - cajainicial(fech + 1)
    'MontoCierre([FechaSistema]![Dates],"C")+MontoCierre([FechaSistema]![Dates],"D")*tipocambio([FechaSistema]![Dates])-cajainicial([FechaSistema]![Dates]+1)
    rst.MoveNext
Next I

rst.Close

Exit Function

Nulo:
AcumuladoCierresDiarios = 0

End Function

Function CantidadVentasGrupoDia(adia As Date, gru As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Anulado, DBBatidos.CodGrupo, Sum([SubPedido]![Cantidad]) AS Cantidad" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, DBBatidos.CodGrupo, NotaDePedido.Fecha" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((DBBatidos.CodGrupo)=" & gru & ") AND ((NotaDePedido.Fecha)=#" & adia & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadVentasGrupoDia = rst("Cantidad")
rst.Close

Exit Function

Nulo:
CantidadVentasGrupoDia = 0

End Function


Function CantidadPedidosValidosDia(adia As Date) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Fecha, NotaDePedido.Anulado, MontoPedido([NotaDePedido]![CodPedido])<>0 AS Monto, Sum(1) AS Total" & _
" FROM NotaDePedido GROUP BY NotaDePedido.Fecha, NotaDePedido.Anulado, MontoPedido([NotaDePedido]![CodPedido])<>0" & _
" HAVING (((NotaDePedido.Fecha)=#" & adia & "#) AND" & _
" ((NotaDePedido.Anulado)=0) AND ((MontoPedido([NotaDePedido]![CodPedido])<>0)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadPedidosValidosDia = rst("Total")
rst.Close

Exit Function

Nulo:
CantidadPedidosValidosDia = 0

End Function

Function CantidadProductosCompradosCajaDia(fechu As Date) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de productos o lineas de compras de caja en un dia
miSQL = "SELECT compras.Fecha, compras.Tipo, Sum(1) AS Total" & _
" FROM compras GROUP BY compras.Fecha, compras.Tipo HAVING (((compras.Fecha)=#" & fechu & "#) AND ((compras.Tipo)='CAJA'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadProductosCompradosCajaDia = rst("Total")
rst.Close

Exit Function

Nulo:
CantidadProductosCompradosCajaDia = 0

End Function

Function TotalRetirosCajaDiaEquivalente(fechu As Date) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Depositos.Fecha, Depositos.DuranteTurno," & _
" Sum(IIf([Depositos]![Denominacion]='Dolares',[Depositos]![Monto]*tipocambio([Depositos]![Fecha]),[Depositos]![Monto])) AS Expr2" & _
" FROM Depositos GROUP BY Depositos.Fecha, Depositos.DuranteTurno" & _
" HAVING (((Depositos.Fecha)=#" & fechu & "#) AND ((Depositos.DuranteTurno)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TotalRetirosCajaDiaEquivalente = rst("Expr2")
rst.Close

Exit Function

Nulo:
TotalRetirosCajaDiaEquivalente = 0

End Function
Function IngresoComprasAutomatico(fechaIn As Date)

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT Compras.CodIngresoAlmacen, Compras.CodCotizacion," & _
        "Compras.Cantidad, Compras.Fecha, Compras.Ingresado " & _
        "FROM Compras " & _
        "WHERE Compras.Fecha = #" & fechaIn & "# AND Compras.Ingresado = False"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (" & rst("CodCotizacion") & ", " & rst("Cantidad") & ", #" & rst("Fecha") & "#)"
    DoCmd.SetWarnings True
    contador = contador + 1
    
    ' Actualizar el estado de ingresado en la tabla Compras
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE Compras SET Ingresado = True WHERE CodIngresoAlmacen = " & rst("CodIngresoAlmacen")
    DoCmd.SetWarnings True
    
    rst.MoveNext
Loop
rst.Close
'MsgBox "Se agregaron " & contador & " productos al registro de ingresos"
Exit Function

Nulo:
MsgBox "Error al crear datos"

End Function
Function ResumenMermasCotiDia(fechi As Date) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT [Merma Cotizacion].Fecha, nombreproductocoti([Merma Cotizacion]![CodCotizacion]) AS produ," & _
" [Merma Cotizacion].Cantidad, Incidencias.Nombre" & _
" FROM [Merma Cotizacion] INNER JOIN Incidencias ON [Merma Cotizacion].CodIncidencia = Incidencias.CodIncidencia" & _
" WHERE ((([Merma Cotizacion].Fecha)=#" & fechi & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
ResumenMermasCotiDia = ""
Do While Not rst.EOF
    ResumenMermasCotiDia = ResumenMermasCotiDia & rst("produ") & " (" & rst("Cantidad") & ") - " & rst("Nombre") & vbCrLf
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
ResumenMermasCotiDia = ""
End Function


Function ResumenMermasIngreDia(fechi As Date) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT [Merma Unidad].Fecha, [DBIngredientes]![Nombre] & ' ' & [DBIngredientes]![Unidad] AS produ," & _
" [Merma Unidad].Cantidad, Incidencias.Nombre" & _
" FROM (Incidencias INNER JOIN [Merma Unidad] ON Incidencias.CodIncidencia = [Merma Unidad].CodIncidencia)" & _
" INNER JOIN DBIngredientes ON [Merma Unidad].CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE ((([Merma Unidad].Fecha)=#" & fechi & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
ResumenMermasIngreDia = ""
Do While Not rst.EOF
    ResumenMermasIngreDia = ResumenMermasIngreDia & rst("produ") & " (" & rst("Cantidad") & ") - " & rst("Nombre") & vbCrLf
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
ResumenMermasIngreDia = ""
End Function


