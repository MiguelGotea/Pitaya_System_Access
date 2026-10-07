' ==========================================================
' Modulo  : Reporte Mensual
' Tipo    : 1
' Lineas  : 924
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database

'---------------------------------------INGRESOS---------------------------------------------------------

Function CierreDiarioxLocal(hor As Date, fech As Date, loc As String) As Double  'GLobal
'CIerre Final de local especifico en fecha especifica a una hora especifica en cordobas

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CierreDiario.Expr1000, CierreDiario.HoraFinal, CierreDiario.Fecha, [CierreDiario]![MFCor]+[CierreDiario]![MFDol]*TipoCambio([CierreDiario]![Fecha]) AS Monto" & _
" FROM CierreDiario IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((CierreDiario.Expr1000)='" & loc & "') AND ((CierreDiario.HoraFinal)=#" & hor & "#) AND ((CierreDiario.Fecha)=#" & fech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CierreDiarioxLocal = rst("Monto")
rst.Close

Exit Function

Nulo:
CierreDiarioxLocal = 0

End Function

Function SumaCierresMesGlobal(M As Integer, a As Integer) As Double  'Global
'Suma de montos de cierre de todos los locales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT Month([CierreDiario]![Fecha]) AS Mes, Year([CierreDiario]![Fecha]) AS Ano, CierreDiario.Fecha, CierreDiario.Expr1000, Max(CierreDiario.HoraFinal) AS MáxDeHoraFinal" & _
" FROM CierreDiario IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([CierreDiario]![Fecha]), Year([CierreDiario]![Fecha]), CierreDiario.Fecha, CierreDiario.Expr1000" & _
" HAVING (((Month([CierreDiario]![Fecha]))=" & M & ") AND ((Year([CierreDiario]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst

Do While Not rst.EOF
    suma = suma + CierreDiarioxLocal(rst("MáxDeHoraFinal"), rst("Fecha"), rst("Expr1000"))
    rst.MoveNext
Loop

rst.Close
SumaCierresMesGlobal = suma

Exit Function

Nulo:
SumaCierresMesGlobal = 0

End Function

Function SumaCierresMesInterno(M As Integer, a As Integer) As Double  'Local
'Suma de montos de cierre de local sistema

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mes, Year([FechaSistema]![Dates]) AS ano, FechaSistema.Dates" & _
" FROM FechaSistema" & _
" WHERE (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    suma = suma + MontoCierre(rst("Dates"), "T")
    rst.MoveNext
Loop

rst.Close

SumaCierresMesInterno = suma

Exit Function

Nulo:
SumaCierresMesInterno = 0

End Function

Function SumaCajaInicialMes(M As Integer, a As Integer) As Double  'Global
'Caja Inicial de todos los locales de un mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Sum(EstadoInicial.Dinero) AS SumaDeDinero, Month([EstadoInicial]![Fecha]) AS Expr1, Year([EstadoInicial]![Fecha]) AS Expr2" & _
" FROM EstadoInicial IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([EstadoInicial]![Fecha]), Year([EstadoInicial]![Fecha])" & _
" HAVING (((Month([EstadoInicial]![Fecha]))=" & M & ") AND ((Year([EstadoInicial]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaCajaInicialMes = rst("SumaDeDinero")
rst.Close

Exit Function

Nulo:
SumaCajaInicialMes = 0

End Function

Function SumaCajaInicialMesInterno(M As Integer, a As Integer) As Double  'Local
'Caja Inicial de todos los locales de un mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([EstadoInicial]![Fecha]) AS mesi, Year([EstadoInicial]![Fecha]) AS ano, Sum(EstadoInicial.Dinero) AS SumaDeDinero" & _
" FROM EstadoInicial" & _
" GROUP BY Month([EstadoInicial]![Fecha]), Year([EstadoInicial]![Fecha])" & _
" HAVING (((Month([EstadoInicial]![Fecha]))=" & M & ") AND ((Year([EstadoInicial]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaCajaInicialMesInterno = rst("SumaDeDinero")
rst.Close

Exit Function

Nulo:
SumaCajaInicialMesInterno = 0

End Function

Function SalidasDeCajaMes(M As Integer, a As Integer) As Double 'Global
'Compras realziadas de caja, vales, de mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Expr1, Year([Compras]![Pagado]) AS Expr2, [Compras]![local]=0 AS conloca," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Compras.Tipo, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), [Compras]![local]=0, Compras.Tipo, IsNull([Compras]![Pagado])" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ")" & _
" AND ((Compras.Tipo)='CAJA') AND ([Compras]![local]=0)=0 AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SalidasDeCajaMes = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SalidasDeCajaMes = 0

End Function

Function SalidasDeCajaMesInterno(M As Integer, a As Integer) As Double 'Local
'Compras realziadas de caja, vales, de mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS mesi, Year([Compras]![Pagado]) AS ano, Compras.Tipo, Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM Compras" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), Compras.Tipo, IsNull([Compras]![Pagado])" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ") AND ((Compras.Tipo)='CAJA') AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SalidasDeCajaMesInterno = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SalidasDeCajaMesInterno = 0

End Function

Function VentasTotalesMesLocalGuardado(M As Integer, a As Integer, loc As Integer) As Double  'Global
'Ventas totoales de cada local segun monto guardado

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT NotadePedido.local, Month([NotaDePedido]![Fecha]) AS ames, Year([NotaDePedido]![Fecha]) AS aano," & _
" Sum(NotadePedido.TotalGuardado) AS SumaDeTotalGuardado" & _
" FROM NotadePedido IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY NotadePedido.local, Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha])" & _
" HAVING (((NotadePedido.local)=" & loc & ") AND ((Month([NotaDePedido]![Fecha]))=" & M & ") AND ((Year([NotaDePedido]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasTotalesMesLocalGuardado = rst("SumaDeTotalGuardado")
rst.Close

Exit Function

Nulo:
VentasTotalesMesLocalGuardado = 0

End Function

Function VentasTotalesMesGuardado(M As Integer, a As Integer) As Double  'local
'Ventas totoales de cada local segun monto guardado

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([NotaDePedido]![Fecha]) AS ames, Year([NotaDePedido]![Fecha]) AS aano," & _
" Sum(NotadePedido.TotalGuardado) AS SumaDeTotalGuardado" & _
" FROM NotadePedido" & _
" GROUP BY Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha])" & _
" HAVING Month([NotaDePedido]![Fecha])=" & M & " AND Year([NotaDePedido]![Fecha])=" & a
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasTotalesMesGuardado = rst("SumaDeTotalGuardado")
rst.Close

Exit Function

Nulo:
VentasTotalesMesGuardado = 0

End Function

Function FRVentaxGrupoMes(M As Integer, a As Integer, gru As Long, Valor As Integer, loc As Integer) As Double 'Global x grupo
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mes, Year([FechaSistema]![Dates]) AS ano, FechaSistema.Dates" & _
" FROM FechaSistema" & _
" WHERE (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

FRVentaxGrupoMes = 0

For I = 1 To Cant
    FRVentaxGrupoMes = FRVentaxGrupoMes + FRVentaxGrupoDia(rst("Dates"), gru, Valor, loc)
    rst.MoveNext
Next I

rst.Close
     
Exit Function

Nulo:
FRVentaxGrupoMes = 0

End Function
Function FRVentaxGrupoMesLocal(M As Integer, a As Integer, gru As Long, Valor As Integer) As Double 'local x grupo
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mes, Year([FechaSistema]![Dates]) AS ano, FechaSistema.Dates" & _
" FROM FechaSistema" & _
" WHERE (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

FRVentaxGrupoMesLocal = 0

For I = 1 To Cant
    FRVentaxGrupoMesLocal = FRVentaxGrupoMesLocal + FRVentaxGrupoDiaInterno(rst("Dates"), gru, Valor)
    rst.MoveNext
Next I

rst.Close
     
Exit Function

Nulo:
FRVentaxGrupoMesLocal = 0

End Function
Function FRVentaTotalMesLocal(M As Integer, a As Integer, Valor As Integer) As Double  'Interno todos grupos
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

FRVentaTotalMesLocal = 0

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mesi, Year([FechaSistema]![Dates]) AS ano, FRVentaxGrupoDiaInterno([FechaSistema]![Dates],[Grupos]![CodGrupo]," & Valor & ") AS Monto" & _
" FROM Grupos, FechaSistema" & _
" WHERE (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
FRVentaTotalMesLocal = FRVentaTotalMesLocal + rst("Monto")
rst.MoveNext
Loop

rst.Close
     
Exit Function

Nulo:
FRVentaTotalMesLocal = 0

End Function

Function FRVentasTotalGruposMes(M As Integer, a As Integer, Valor As Integer) As Double  'GLobal todos grupos
'Suma de todos los grupos de FRVentaxGrupoMes 
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Sum(FRVentaxGrupoMes(" & M & "," & a & ",[Grupos]![CodGrupo]," & Valor & ",[DatosSistema]![CodSistema])) AS Monto" & _
" FROM Grupos, DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRVentasTotalGruposMes = rst("Monto")
rst.Close

Exit Function

Nulo:
FRVentasTotalGruposMes = 0

End Function


Function FRVentaTotalMesTipoDeliv(M As Integer, a As Integer, adeli As Integer, atipo As Integer, montocant As Integer, loca As Integer) As Double  'Global todos grupos x tipopago y x deliv
'montocant : 1 para monto, 2 para cantidad

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

FRVentaTotalMesTipoDeliv = 0

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mesi, Year([FechaSistema]![Dates]) AS ano," & _
" FRVentaxDiaTipoDeliv([FechaSistema]![Dates], " & atipo & ", " & adeli & ", 0, 0, " & montocant & ", " & loca & ") AS Monto" & _
" FROM FechaSistema" & _
" WHERE (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
FRVentaTotalMesTipoDeliv = FRVentaTotalMesTipoDeliv + rst("Monto")
rst.MoveNext
Loop

rst.Close
Exit Function

Nulo:
FRVentaTotalMesTipoDeliv = 0

End Function

Function FRVentaTotalMesTipoDelivGuardado(M As Integer, a As Integer, adeli As Integer, atipo As Integer, loca As Integer) As Double  'Global todos grupos x tipopago y x deliv
'Venta totoal montos

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([FechaSistema]![Dates]) AS ames, Year([FechaSistema]![Dates]) AS aano," & _
" NotadePedido.Delivery, NotadePedido.POS, NotadePedido.local, Sum(NotadePedido.TotalGuardado) AS SumaDeTotalGuardado" & _
" FROM NotadePedido INNER JOIN FechaSistema ON NotadePedido.Fecha = FechaSistema.Dates" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([FechaSistema]![Dates]), Year([FechaSistema]![Dates]), NotadePedido.Delivery," & _
" NotadePedido.POS, NotadePedido.local" & _
" HAVING (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & ")" & _
" AND ((NotadePedido.Delivery)=" & adeli & ") AND ((NotadePedido.POS)=" & atipo & ") AND ((NotadePedido.local)=" & loca & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRVentaTotalMesTipoDelivGuardado = rst("SumaDeTotalGuardado")
rst.Close
Exit Function

Nulo:
FRVentaTotalMesTipoDelivGuardado = 0

End Function

Function FRVentaInternoMesTipoDelivGuardado(M As Integer, a As Integer, adeli As Integer, atipo As Integer) As Double  'Interno todos grupos x tipopago y x deliv
'Venta totoal montos Interno

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([FechaSistema]![Dates]) AS ames, Year([FechaSistema]![Dates]) AS aano," & _
" NotadePedido.Delivery, NotadePedido.POS, Sum(NotadePedido.TotalGuardado) AS SumaDeTotalGuardado" & _
" FROM NotadePedido INNER JOIN FechaSistema ON NotadePedido.Fecha = FechaSistema.Dates" & _
" GROUP BY Month([FechaSistema]![Dates]), Year([FechaSistema]![Dates]), NotadePedido.Delivery," & _
" NotadePedido.POS" & _
" HAVING (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & ")" & _
" AND ((NotadePedido.Delivery)=" & adeli & ") AND ((NotadePedido.POS)=" & atipo & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRVentaInternoMesTipoDelivGuardado = rst("SumaDeTotalGuardado")
rst.Close
Exit Function

Nulo:
FRVentaInternoMesTipoDelivGuardado = 0

End Function

Function FRVentaTotalMesGrupoTipoDeliv(M As Integer, a As Integer, agru As Integer, adeli As Integer, atipo As Integer, montocant As Integer, loca As Integer) As Double  'Global todos grupos x tipopago y x deliv
'montocant : 1 para monto, 2 para cantidad

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

FRVentaTotalMesGrupoTipoDeliv = 0

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mesi, Year([FechaSistema]![Dates]) AS ano, Grupos.CodGrupo" & _
" FRVentaxGrupoDiaTipoDeliv([FechaSistema]![Dates], " & agru & ", " & atipo & ", " & adeli & ", 0, 0, " & montocant & ", " & loca & ") AS Monto" & _
" FROM fechasistema, Grupos" & _
" WHERE Month([FechaSistema]![Dates])=" & M & " AND Year([FechaSistema]![Dates])=" & a & " AND Grupos.CodGrupo=" & agru
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
FRVentaTotalMesGrupoTipoDeliv = FRVentaTotalMesGrupoTipoDeliv + rst("Monto")
rst.MoveNext
Loop

rst.Close
Exit Function

Nulo:
FRVentaTotalMesGrupoTipoDeliv = 0

End Function

Function FRVentaTotalMesGrupo(M As Integer, a As Integer, agru As Integer, montocant As Integer, loca As Integer) As Double  'Global todos grupos x tipopago y x deliv
'montocant : 1 para monto, 2 para cantidad

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

FRVentaTotalMesGrupo = 0

miSQL = "SELECT Month([FechaSistema]![Dates]) AS mesi, Year([FechaSistema]![Dates]) AS ano, Grupos.CodGrupo," & _
" FRVentaxGrupoDiaTipoDeliv([FechaSistema]![Dates], " & agru & ", [TipoPago]![CodTipo], [Delivery]![CodDelivery], 0, 0, " & montocant & ", " & loca & ") AS Monto" & _
" FROM fechasistema, Grupos, Delivery, TipoPago" & _
" WHERE (((Month([FechaSistema]![Dates]))=" & M & ") AND ((Year([FechaSistema]![Dates]))=" & a & ") AND ((Grupos.CodGrupo)=" & agru & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
FRVentaTotalMesGrupo = FRVentaTotalMesGrupo + rst("Monto")
rst.MoveNext
Loop

rst.Close
Exit Function

Nulo:
FRVentaTotalMesGrupo = 0

End Function

'---------------------------------------EGRESOS-----------------------------------------------------------

Function ComprasGrupoMes(M As Integer, a As Integer, TIPO1 As String, TIPO2 As String) As Double  'global
'Compras realizadas de un grupo especific en mes y ano especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Expr1, Year([Compras]![Pagado]) AS Expr2, DBIngredientes.TIPO1," & _
" DBIngredientes.TIPO2, Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM (Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion)" & _
" INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), DBIngredientes.TIPO1, DBIngredientes.TIPO2," & _
" IsNull([Compras]![Pagado])" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ")" & _
" AND ((DBIngredientes.TIPO1)='" & TIPO1 & "') AND ((DBIngredientes.TIPO2)='" & TIPO2 & "') AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

ComprasGrupoMes = rst("SumaDeCostoTotal")  '- ComprasGlobalMesAlmacenGlobal(M, A)
'compra de porciones alamacen global no aplica por ahora
Exit Function

Nulo:
'MsgBox "error"
ComprasGrupoMes = 0

End Function
Function ComprasGrupoMesLocal(M As Integer, a As Integer, TIPO1 As String, TIPO2 As String, loc As Integer) As Double  'global
'Compras realizadas de un grupo especific en mes y ano especifico por local

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Expr1, Year([Compras]![Pagado]) AS Expr2, DBIngredientes.TIPO1," & _
" DBIngredientes.TIPO2, Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo, Compras.local" & _
" FROM (Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion)" & _
" INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), DBIngredientes.TIPO1, DBIngredientes.TIPO2," & _
" IsNull([Compras]![Pagado]), Compras.local" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ")" & _
" AND ((DBIngredientes.TIPO1)='" & TIPO1 & "') AND ((DBIngredientes.TIPO2)='" & TIPO2 & "') AND ((IsNull([Compras]![Pagado]))=0) AND ((Compras.local)=" & loc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

ComprasGrupoMesLocal = rst("SumaDeCostoTotal")  '- ComprasGlobalMesAlmacenGlobal(M, A)
'compra de porciones alamacen global no aplica por ahora
Exit Function

Nulo:
'MsgBox "error"
ComprasGrupoMesLocal = 0

End Function
Function ComprasGrupoMesNoControl(M As Integer, a As Integer, TIPO1 As String, TIPO2 As String) As Double  'local
'Compras realizadas de un grupo especific en mes y ano especifico
'Administraticos osea no control

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Fecha]) AS Expr1, Year([Compras]![Fecha]) AS Expr2, DBIngredientes.TIPO1, DBIngredientes.TIPO2, Sum(Compras.CostoTotal) AS SumaDeCostoTotal" & _
" FROM (Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion) INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" GROUP BY Month([Compras]![Fecha]), Year([Compras]![Fecha]), DBIngredientes.TIPO1, DBIngredientes.TIPO2" & _
" HAVING (((Month([Compras]![Fecha]))=" & M & ") AND ((Year([Compras]![Fecha]))=" & a & ") AND ((DBIngredientes.TIPO1)='" & TIPO1 & "') AND ((DBIngredientes.TIPO2)='" & TIPO2 & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasGrupoMesNoControl = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
ComprasGrupoMesNoControl = 0

End Function

Function ComprasGrupoMesControl(M As Integer, a As Integer, TIPO1 As String, TIPO2 As String) As Double  'local
'Compras realizadas de un grupo especific en mes y ano especifico
'Administraticos osea no control

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT Month([IngresosPitaya]![Fecha]) AS Expr1, Year([IngresosPitaya]![Fecha]) AS Expr2, DBIngredientes.TIPO1, DBIngredientes.TIPO2, numerosemana([IngresosPitaya]![Fecha]) AS semana, IngresosPitaya.CodCotizacion, Sum(IngresosPitaya.Cantidad) AS SumaDeCantidad" & _
" FROM (Cotizaciones INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion" & _
" GROUP BY Month([IngresosPitaya]![Fecha]), Year([IngresosPitaya]![Fecha]), DBIngredientes.TIPO1, DBIngredientes.TIPO2, numerosemana([IngresosPitaya]![Fecha]), IngresosPitaya.CodCotizacion" & _
" HAVING (((Month([IngresosPitaya]![Fecha]))=" & M & ") AND ((Year([IngresosPitaya]![Fecha]))=" & a & ") AND ((DBIngredientes.TIPO1)='" & TIPO1 & "') AND ((DBIngredientes.TIPO2)='" & TIPO2 & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("SumaDeCantidad") * CostoUnitarioUnidad(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop

rst.Close
ComprasGrupoMesControl = suma

Exit Function

Nulo:
ComprasGrupoMesControl = 0

End Function

Function ComprasGlobalMes(M As Integer, a As Integer) As Double
'Compras realizadas en mes y ano especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Expr1, Year([Compras]![Pagado]) AS Expr2," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), IsNull([Compras]![Pagado])" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ") AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasGlobalMes = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
ComprasGlobalMes = 0

End Function

Function ComprasGlobalSistema0(M As Integer, a As Integer) As Double
'Compras realizadas en mes y ano especifico del sistema 0

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Expr1, Year([Compras]![Pagado]) AS Expr2," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya0_DB.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), IsNull([Compras]![Pagado])" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ") AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasGlobalSistema0 = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
ComprasGlobalSistema0 = 0

End Function

Function ComprasGrupoMesGlobalSistema0(M As Integer, a As Integer, TIPO1 As String, TIPO2 As String) As Double  'sistema0
'Compras realizadas de un grupo especific en mes y ano especifico del sistema0

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Expr1, Year([Compras]![Pagado]) AS Expr2, DBIngredientes.TIPO1," & _
" DBIngredientes.TIPO2, Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM (Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion)" & _
" INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya0_System.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), DBIngredientes.TIPO1, DBIngredientes.TIPO2," & _
" IsNull([Compras]![Pagado])" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ")" & _
" AND ((DBIngredientes.TIPO1)='" & TIPO1 & "') AND ((DBIngredientes.TIPO2)='" & TIPO2 & "') AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

ComprasGrupoMesGlobalSistema0 = rst("SumaDeCostoTotal")

Exit Function

Nulo:
ComprasGrupoMesGlobalSistema0 = 0

End Function

Function ComprasGlobalMesAlmacenGlobal(M As Integer, a As Integer) As Double
'Compras realizadas en mes y ano especifico de productos almacen global

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Month([Compras]![Pagado]) AS Me, Year([Compras]![Pagado]) AS An," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS Nulo, Cotizaciones.Marca" & _
" FROM Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([Compras]![Pagado]), Year([Compras]![Pagado]), IsNull([Compras]![Pagado]), Cotizaciones.Marca" & _
" HAVING (((Month([Compras]![Pagado]))=" & M & ") AND ((Year([Compras]![Pagado]))=" & a & ")" & _
" AND ((IsNull([Compras]![Pagado]))=0) AND ((Cotizaciones.Marca)='Almacen Global'));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasGlobalMesAlmacenGlobal = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
ComprasGlobalMesAlmacenGlobal = 0

End Function


Function FRCostoConsumoTeoricoMes(M As Integer, a As Integer) As Double 'local
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0
'Suma de costo de cantidades consumidas teoricos*
miSQL = "SELECT Month([NotaDePedido]![Fecha]) AS mesi, Year([NotaDePedido]![Fecha]) AS ano," & _
" NotaDePedido.Anulado, SubReceta.CodIngrediente, numerosemana([NotaDePedido]![Fecha]) AS semana," & _
" Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Cantidad" & _
" FROM (SubPedido INNER JOIN (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido" & _
" GROUP BY Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha]), NotaDePedido.Anulado," & _
" SubReceta.CodIngrediente, numerosemana([NotaDePedido]![Fecha])" & _
" HAVING (((Month([NotaDePedido]![Fecha]))=" & M & ") AND ((Year([NotaDePedido]![Fecha]))=" & a & ")" & _
" AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioIngrediente(rst("CodIngrediente"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoConsumoTeoricoMes = suma
Exit Function

Nulo:
FRCostoConsumoTeoricoMes = 0

End Function

Function FRCostoIngresosMes(M As Integer, a As Integer) As Double 'local
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT DBIngredientes.TIPO1, Month([IngresosPitaya]![Fecha]) AS mesi, Year([IngresosPitaya]![Fecha]) AS ano, IngresosPitaya.CodCotizacion, numerosemana([IngresosPitaya]![Fecha]) AS semana, Sum([IngresosPitaya]![Cantidad]) AS Cantidad" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, Month([IngresosPitaya]![Fecha]), Year([IngresosPitaya]![Fecha]), IngresosPitaya.CodCotizacion, numerosemana([IngresosPitaya]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((Month([IngresosPitaya]![Fecha]))=" & M & ") AND ((Year([IngresosPitaya]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioCotizacion(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoIngresosMes = suma
Exit Function

Nulo:
FRCostoIngresosMes = 0

End Function

Function FRCostoMermasCotizacionMes(M As Integer, a As Integer) As Double 'local
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT DBIngredientes.TIPO1, Month([Merma Cotizacion]![Fecha]) AS mesi, Year([Merma Cotizacion]![Fecha]) AS ano, [Merma Cotizacion].CodCotizacion, numerosemana([Merma Cotizacion]![Fecha]) AS semana, Sum([Merma Cotizacion]![Cantidad]) AS Cantidad" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN [Merma Cotizacion] ON Cotizaciones.CodCotizacion = [Merma Cotizacion].CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, Month([Merma Cotizacion]![Fecha]), Year([Merma Cotizacion]![Fecha]), [Merma Cotizacion].CodCotizacion, numerosemana([Merma Cotizacion]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((Month([Merma Cotizacion]![Fecha]))=" & M & ") AND ((Year([Merma Cotizacion]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioCotizacion(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoMermasCotizacionMes = suma
Exit Function

Nulo:
FRCostoMermasCotizacionMes = 0

End Function

Function FRCostoMermasIngredienteMes(M As Integer, a As Integer) As Double 'local
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT DBIngredientes.TIPO1, Month([Merma Unidad]![Fecha]) AS mesi, Year([Merma Unidad]![Fecha]) AS ano, [Merma Unidad].CodIngrediente, numerosemana([Merma Unidad]![Fecha]) AS semana, Sum([Merma Unidad]![Cantidad]) AS Cantidad" & _
" FROM DBIngredientes INNER JOIN [Merma Unidad] ON DBIngredientes.CodIngrediente = [Merma Unidad].CodIngrediente" & _
" GROUP BY DBIngredientes.TIPO1, Month([Merma Unidad]![Fecha]), Year([Merma Unidad]![Fecha]), [Merma Unidad].CodIngrediente, numerosemana([Merma Unidad]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((Month([Merma Unidad]![Fecha]))=" & M & ") AND ((Year([Merma Unidad]![Fecha]))=" & a & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioIngrediente(rst("CodIngrediente"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoMermasIngredienteMes = suma
Exit Function

Nulo:
FRCostoMermasIngredienteMes = 0

End Function

'--------------------------------------------------------------------------------------------------------
'-------------------------------------costos intervalos--------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'Inventario Final del Mes = Inventario ultimo domingo + Ingresos Intervalo Final - Consumos Intervalo Final - Mermas Intervalo Final

Function FRCostoIngresoIntervalo(desde As Date, hasta As Date) As Double 'local
'suma de1 costo de ingresos de un intvalo de fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0
miSQL = "SELECT DBIngredientes.TIPO1, IngresosPitaya.Fecha, IngresosPitaya.CodCotizacion, numerosemana([IngresosPitaya]![Fecha]) AS semana, Sum(IngresosPitaya.Cantidad) AS SumaDeCantidad" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, IngresosPitaya.Fecha, IngresosPitaya.CodCotizacion, numerosemana([IngresosPitaya]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((IngresosPitaya.Fecha) Between #" & desde & "# And #" & hasta & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("SumaDeCantidad") * FRCostoUnitarioCotizacion(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoIngresoIntervalo = suma
Exit Function

Nulo:
FRCostoIngresoIntervalo = 0

End Function

Function FRCostoConsumoIntervalo(desde As Date, hasta As Date) As Double 'local
'suma de1 costo de consumos de un intvalo de fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0
miSQL = "SELECT NotaDePedido.Fecha, NotaDePedido.Anulado, SubReceta.CodIngrediente," & _
" numerosemana([NotaDePedido]![Fecha]) AS semana," & _
" Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Cantidad" & _
" FROM NotaDePedido INNER JOIN (SubPedido INNER JOIN (DBBatidos INNER JOIN SubReceta" & _
" ON DBBatidos.CodBatido = SubReceta.CodBatido) ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY NotaDePedido.Fecha, NotaDePedido.Anulado, SubReceta.CodIngrediente, numerosemana([NotaDePedido]![Fecha])" & _
" HAVING (((NotaDePedido.Fecha) Between #" & desde & "# And #" & hasta & "#) AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioIngrediente(rst("CodIngrediente"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoConsumoIntervalo = suma
Exit Function

Nulo:
FRCostoConsumoIntervalo = 0

End Function

Function FRCostoMermaIngredienteIntervalo(desde As Date, hasta As Date) As Double 'local
'suma de1 costo de merma de ingrediente de un intvalo de fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0
miSQL = "SELECT DBIngredientes.TIPO1, [Merma Unidad].Fecha, [Merma Unidad].CodIngrediente, numerosemana([Merma Unidad]![Fecha]) AS semana, Sum([Merma Unidad].Cantidad) AS SumaDeCantidad" & _
" FROM [Merma Unidad] INNER JOIN DBIngredientes ON [Merma Unidad].CodIngrediente = DBIngredientes.CodIngrediente" & _
" GROUP BY DBIngredientes.TIPO1, [Merma Unidad].Fecha, [Merma Unidad].CodIngrediente, numerosemana([Merma Unidad]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND (([Merma Unidad].Fecha) Between #" & desde & "# And #" & hasta & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("SumaDeCantidad") * FRCostoUnitarioIngrediente(rst("CodIngrediente"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoMermaIngredienteIntervalo = suma
Exit Function

Nulo:
FRCostoMermaIngredienteIntervalo = 0

End Function

Function FRCostoMermaCotizacionIntervalo(desde As Date, hasta As Date) As Double 'local
'suma de1 costo de merma de cotizacion de un intvalo de fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0
miSQL = "SELECT DBIngredientes.TIPO1, [Merma Cotizacion].Fecha, [Merma Cotizacion].CodCotizacion, numerosemana([Merma Cotizacion]![Fecha]) AS semana, Sum([Merma Cotizacion].Cantidad) AS SumaDeCantidad" & _
" FROM [Merma Cotizacion] INNER JOIN (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) ON [Merma Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, [Merma Cotizacion].Fecha, [Merma Cotizacion].CodCotizacion, numerosemana([Merma Cotizacion]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND (([Merma Cotizacion].Fecha) Between #" & desde & "# And #" & hasta & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("SumaDeCantidad") * FRValorCotizacion(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoMermaCotizacionIntervalo = suma
Exit Function

Nulo:
FRCostoMermaCotizacionIntervalo = 0

End Function

'-----------------------------------------------------------------------------------------------------------------
