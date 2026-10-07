' ==========================================================
' Modulo  : Reporte Semanal
' Tipo    : 1
' Lineas  : 1940
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database

Function SumaCierresSemanaInterno(sem As Integer) As Double  'Local
'Suma de montos de cierre de local sistema

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT numerosemana([FechaSistema]![Dates]) AS semana, FechaSistema.Dates" & _
" FROM FechaSistema" & _
" WHERE (((numerosemana([FechaSistema]![Dates]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    suma = suma + MontoCierre(rst("Dates"), "T")
    rst.MoveNext
Loop

rst.Close

SumaCierresSemanaInterno = suma

Exit Function

Nulo:
SumaCierresSemanaInterno = 0

End Function

Function SumaCajaInicialSemanaInterno(sem As Integer) As Double  'Local
'Caja Inicial de todos los locales de un semana especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([EstadoInicial]![Fecha]) AS semana, Sum(EstadoInicial.Dinero) AS SumaDeDinero" & _
" FROM EstadoInicial" & _
" GROUP BY numerosemana([EstadoInicial]![Fecha])" & _
" HAVING ((numerosemana([EstadoInicial]![Fecha]))=" & sem & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaCajaInicialSemanaInterno = rst("SumaDeDinero")
rst.Close

Exit Function

Nulo:
SumaCajaInicialSemanaInterno = 0

End Function

Function SalidasDeCajaSemanaInterno(sem As Integer) As Double 'Local
'Compras realziadas de caja, vales, de semana especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([Compras]![Pagado]) AS semana, Compras.Tipo, Sum(Compras.CostoTotal) AS SumaDeCostoTotal, IsNull([Compras]![Pagado]) AS nulo" & _
" FROM Compras" & _
" GROUP BY numerosemana([Compras]![Pagado]), Compras.Tipo, IsNull([Compras]![Pagado])" & _
" HAVING (((numerosemana([Compras]![Pagado]))=" & sem & ") AND ((Compras.Tipo)='CAJA') AND ((IsNull([Compras]![Pagado]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SalidasDeCajaSemanaInterno = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SalidasDeCajaSemanaInterno = 0

End Function

Function FRVentaTotalSemana(sem As Integer, Valor As Integer, loci As Integer) As Double  'Global
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS
'loci = local a buscar

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

FRVentaTotalSemana = 0

miSQL = "SELECT numerosemana([FechaSistema]![Dates]) AS semana, FRVentaxGrupoDia([FechaSistema]![Dates],[Grupos]![CodGrupo]," & Valor & ", " & loci & ") AS Monto" & _
" FROM Grupos, FechaSistema" & _
" WHERE ((numerosemana([FechaSistema]![Dates]))=" & sem & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
FRVentaTotalSemana = FRVentaTotalSemana + rst("Monto")
rst.MoveNext
Loop

rst.Close
     
Exit Function

Nulo:
FRVentaTotalSemana = 0

End Function


Function FRVentaTotalSemanaLocal(sem As Integer, Valor As Integer) As Double  'Interno
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

FRVentaTotalSemanaLocal = 0

miSQL = "SELECT numerosemana([FechaSistema]![Dates]) AS semana, FRVentaxGrupoDiaInterno([FechaSistema]![Dates],[Grupos]![CodGrupo]," & Valor & ") AS Monto" & _
" FROM Grupos, FechaSistema" & _
" WHERE ((numerosemana([FechaSistema]![Dates]))=" & sem & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
FRVentaTotalSemanaLocal = FRVentaTotalSemanaLocal + rst("Monto")
rst.MoveNext
Loop

rst.Close
     
Exit Function

Nulo:
FRVentaTotalSemanaLocal = 0

End Function

Function FRCostoUnitarioCotizacion(coti As Integer, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalCUCotiSemanal.Semana, CalCUCotiSemanal.CodCotizacion, CalCUCotiSemanal.CU, CalCUCotiSemanal.FechaPublicada" & _
" FROM CalCUCotiSemanal" & _
" WHERE (((CalCUCotiSemanal.semana) = " & sem & ") And ((CalCUCotiSemanal.CodCotizacion) = " & coti & "))" & _
" ORDER BY CalCUCotiSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRCostoUnitarioCotizacion = rst("CU")
rst.Close
     
Exit Function

Nulo:
FRCostoUnitarioCotizacion = 0

End Function

Function FRValorCotizacion(coti As Integer, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalCUCotiSemanal.Semana, CalCUCotiSemanal.CodCotizacion, CalCUCotiSemanal.Valor, CalCUCotiSemanal.FechaPublicada" & _
" FROM CalCUCotiSemanal" & _
" WHERE (((CalCUCotiSemanal.semana) = " & sem & ") And ((CalCUCotiSemanal.CodCotizacion) = " & coti & "))" & _
" ORDER BY CalCUCotiSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRValorCotizacion = rst("Valor")
rst.Close
     
Exit Function

Nulo:
FRValorCotizacion = 0

End Function

Function FRCostoUnitarioIngrediente(ING As String, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalCUIngSemanal.CodIngrediente, CalCUIngSemanal.CU, CalCUIngSemanal.Semana, CalCUIngSemanal.FechaPublicada" & _
" FROM CalCUIngSemanal" & _
" WHERE (((CalCUIngSemanal.CodIngrediente) = '" & ING & "') And ((CalCUIngSemanal.semana) = " & sem & "))" & _
" ORDER BY CalCUIngSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRCostoUnitarioIngrediente = rst("CU")
rst.Close
     
Exit Function

Nulo:
FRCostoUnitarioIngrediente = 0

End Function

Function FRConversionCotizacion(coti As Integer, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConversionSemanal.CodCotizacion, CalConversionSemanal.Conversion, CalConversionSemanal.Semana, CalConversionSemanal.FechaPublicada" & _
" FROM CalConversionSemanal" & _
" WHERE (((CalConversionSemanal.CodCotizacion) = " & coti & ") And ((CalConversionSemanal.semana) = " & sem & "))" & _
" ORDER BY CalConversionSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConversionCotizacion = rst("Conversion")
rst.Close
     
Exit Function

Nulo:
FRConversionCotizacion = 0

End Function
Function FRConversionCotizacionLocal(coti As Integer, sem As Integer, loci As Integer) As Double
'Conversion calculada de cada local gaurdado
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConversionSemanal.CodCotizacion, CalConversionSemanal.Conversion, CalConversionSemanal.Semana," & _
" CalConversionSemanal.FechaPublicada, CalConversionSemanal.local" & _
" FROM CalConversionSemanal IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((CalConversionSemanal.CodCotizacion) = " & coti & ") And ((CalConversionSemanal.semana) = " & sem & ")" & _
" And ((CalConversionSemanal.local) = " & loci & "))" & _
" ORDER BY CalConversionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConversionCotizacionLocal = rst("Conversion")
rst.Close
     
Exit Function

Nulo:
FRConversionCotizacionLocal = 0

End Function

Function FRConsumoIngrediente(inge As String, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, Consumo, Semana, FechaPublicada" & _
" FROM CalConsumoIngSemanal" & _
" WHERE (((CodIngrediente) = '" & inge & "') And ((semana) = " & sem & "))" & _
" ORDER BY FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoIngrediente = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoIngrediente = 0

End Function


Function FRConsumoIngredienteSumaLocales(ingresl As String, semasl As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE ((StatusSucursales.Sucursal<>0) AND ((StatusSucursales.Activo)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoIngredienteSumaLocales = 0
Do While Not rst.EOF
    FRConsumoIngredienteSumaLocales = FRConsumoIngredienteSumaLocales + FRConsumoIngredienteSucursalArchivoAdjunto(ingresl, semasl, rst("CodLocal"))
    rst.MoveNext
Loop
rst.Close
Exit Function

Nulo:
FRConsumoIngredienteSumaLocales = 0
End Function

Function FRConsumoIngredienteSucursalArchivoAdjunto(ingred As String, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, Consumo, Semana," & _
" FechaPublicada" & _
" FROM CalConsumoIngSemanal" & loc & _
" WHERE (((CodIngrediente) = '" & ingred & "') And ((semana) = " & sem & "))" & _
" ORDER BY FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoIngredienteSucursalArchivoAdjunto = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoIngredienteSucursalArchivoAdjunto = 0

End Function

Function FRConsumoMaximoCalculadoIngrediente(inge As String, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoIngSemanal.CodIngrediente, CalConsumoIngSemanal.ConsumoMaximo, CalConsumoIngSemanal.Semana, CalConsumoIngSemanal.FechaPublicada" & _
" FROM CalConsumoIngSemanal" & _
" WHERE (((CalConsumoIngSemanal.CodIngrediente) = '" & inge & "') And ((CalConsumoIngSemanal.semana) = " & sem & "))" & _
" ORDER BY CalConsumoIngSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoMaximoCalculadoIngrediente = rst("ConsumoMaximo")
rst.Close
     
Exit Function

Nulo:
FRConsumoMaximoCalculadoIngrediente = 0

End Function

Function FRConsumoSemanalMaximoIngrediente(ingred As String, sem As Integer, rango As Integer) As Double
Dim cont As Integer
Dim temp As Double

cont = 0
temp = FRConsumoIngrediente(ingred, sem - 1 - cont)

Do While cont < rango
If FRConsumoIngrediente(ingred, sem - 1 - cont) > temp Then
    temp = FRConsumoIngrediente(ingred, sem - 1 - cont)
End If
cont = cont + 1
Loop
FRConsumoSemanalMaximoIngrediente = temp
Exit Function

Nulo:
FRConsumoSemanalMaximoIngrediente = 0

End Function

Function FRConsumoNoPorcion(ingred As String, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoNoPorcionSemanal.CodIngrediente, CalConsumoNoPorcionSemanal.Consumo, CalConsumoNoPorcionSemanal.Semana," & _
" CalConsumoNoPorcionSemanal.FechaPublicada" & _
" FROM CalConsumoNoPorcionSemanal" & _
" WHERE (((CalConsumoNoPorcionSemanal.CodIngrediente) = '" & ingred & "') And ((CalConsumoNoPorcionSemanal.semana) = " & sem & "))" & _
" ORDER BY CalConsumoNoPorcionSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoNoPorcion = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoNoPorcion = 0

End Function

Function FRConsumoNoPorcionSucursalArchivoAdjunto(ingred As String, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, Consumo, Semana," & _
" FechaPublicada" & _
" FROM CalConsumoNoPorcionSemanal" & loc & _
" WHERE (((CodIngrediente) = '" & ingred & "') And ((semana) = " & sem & "))" & _
" ORDER BY FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoNoPorcionSucursalArchivoAdjunto = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoNoPorcionSucursalArchivoAdjunto = 0

End Function

Function FRConsumoNoPorcionSumaLocales(ingresl As String, semasl As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE ((StatusSucursales.Sucursal<>0) AND ((StatusSucursales.Activo)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoNoPorcionSumaLocales = 0
Do While Not rst.EOF
    FRConsumoNoPorcionSumaLocales = FRConsumoNoPorcionSumaLocales + FRConsumoNoPorcionSucursalArchivoAdjunto(ingresl, semasl, rst("CodLocal"))
    rst.MoveNext
Loop
rst.Close
Exit Function

Nulo:
FRConsumoNoPorcionSumaLocales = 0
End Function


Function FRConsumoMaximoNoPorcion(ingred As String, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoNoPorcionSemanal.CodIngrediente, CalConsumoNoPorcionSemanal.ConsumoMaximo, CalConsumoNoPorcionSemanal.Semana," & _
" CalConsumoNoPorcionSemanal.FechaPublicada" & _
" FROM CalConsumoNoPorcionSemanal" & _
" WHERE (((CalConsumoNoPorcionSemanal.CodIngrediente) = '" & ingred & "') And ((CalConsumoNoPorcionSemanal.semana) = " & sem & "))" & _
" ORDER BY CalConsumoNoPorcionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoMaximoNoPorcion = rst("ConsumoMaximo")
rst.Close
     
Exit Function

Nulo:
FRConsumoMaximoNoPorcion = 0

End Function

Function FRConsumoNoPorcionxLocal(ingrex As String, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoNoPorcionSemanal.CodIngrediente, CalConsumoNoPorcionSemanal.Consumo," & _
" CalConsumoNoPorcionSemanal.Semana, CalConsumoNoPorcionSemanal.FechaPublicada, CalConsumoNoPorcionSemanal.local" & _
" FROM CalConsumoNoPorcionSemanal IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (CalConsumoNoPorcionSemanal.CodIngrediente = '" & ingrex & "' And CalConsumoNoPorcionSemanal.semana = " & sem & "" & _
" And CalConsumoNoPorcionSemanal.local = " & loc & ")" & _
" ORDER BY CalConsumoNoPorcionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoNoPorcionxLocal = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoNoPorcionxLocal = 0

End Function

Function FRConsumoMaximoNoPorcionxLocal(ingrex As String, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoNoPorcionSemanal.CodIngrediente, CalConsumoNoPorcionSemanal.ConsumoMaximo," & _
" CalConsumoNoPorcionSemanal.Semana, CalConsumoNoPorcionSemanal.FechaPublicada, CalConsumoNoPorcionSemanal.local" & _
" FROM CalConsumoNoPorcionSemanal IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (CalConsumoNoPorcionSemanal.CodIngrediente = '" & ingrex & "' And CalConsumoNoPorcionSemanal.semana = " & sem & "" & _
" And CalConsumoNoPorcionSemanal.local = " & loc & ")" & _
" ORDER BY CalConsumoNoPorcionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoMaximoNoPorcionxLocal = rst("ConsumoMaximo")
rst.Close
     
Exit Function

Nulo:
FRConsumoMaximoNoPorcionxLocal = 0

End Function

Function FRRequerimientoMaximoNoPorciones(ingred As String, sem As Integer, rango As Integer) As Double
On Error GoTo Nulo
Dim cont As Integer
Dim temp As Double

cont = 0
temp = FRConsumoNoPorcion(ingred, sem - 1 - cont)

Do While cont < rango
If FRConsumoNoPorcion(ingred, sem - 1 - cont) > temp Then
    temp = FRConsumoNoPorcion(ingred, sem - 1 - cont)
End If
cont = cont + 1
Loop
FRRequerimientoMaximoNoPorciones = temp
Exit Function

Nulo:
FRRequerimientoMaximoNoPorciones = 0

End Function

Function FRRequerimientoMaximoNoPorcionesxLocalCalculado(ingret As String, sem As Integer, rango As Integer, loc As Integer) As Double
On Error GoTo Nulo


If rango = 4 Then ' jaala directo de la db
    FRRequerimientoMaximoNoPorcionesxLocalCalculado = FRConsumoMaximoNoPorcionxLocal(ingret, sem - 1, loc)
Else
    FRRequerimientoMaximoNoPorcionesxLocalCalculado = FRRequerimientoMaximoNoPorcionesxLocal(ingret, sem, rango, loc)
End If

Exit Function

Nulo:
FRRequerimientoMaximoNoPorcionesxLocalCalculado = 0

End Function

Function FRRequerimientoMaximoNoPorcionesxLocal(ingret As String, sem As Integer, rango As Integer, loc As Integer) As Double
On Error GoTo Nulo
Dim cont As Integer
Dim temp As Double
Dim temp2 As Double

cont = 0
temp = FRConsumoNoPorcionxLocal(ingret, sem - 1 - cont, loc)

Do While cont < rango
    temp2 = FRConsumoNoPorcionxLocal(ingret, sem - 1 - cont, loc)
    If temp2 > temp Then
        temp = temp2
    End If
    cont = cont + 1
Loop
FRRequerimientoMaximoNoPorcionesxLocal = temp
Exit Function

Nulo:
FRRequerimientoMaximoNoPorcionesxLocal = 0

End Function



Function FRRequerimientoMaximoxDespachoNoPorcionesxLocal(ingret As String, sem As Integer, rango As Integer, loc As Integer) As Double
On Error GoTo Nulo
Dim cont As Integer
Dim temp As Double
Dim temp2 As Double
Dim paque As Double


cont = 0
temp = FRConsumoNoPorcionxLocal(ingret, sem - 1 - cont, loc)
paque = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & CotiPrincipalDeIngrediente(ingret))

Do While cont < rango
    temp2 = redondear_mas(FRConsumoNoPorcionxLocal(ingret, sem - 1 - cont, loc) / paque) * paque
    If temp2 > temp Then
        temp = temp2
    End If
    cont = cont + 1
Loop
FRRequerimientoMaximoxDespachoNoPorcionesxLocal = temp
Exit Function

Nulo:
FRRequerimientoMaximoxDespachoNoPorcionesxLocal = 0

End Function
Function FRRequerimientoMaximoxDespachoNoPorcionesxLocalCalculado(ingret As String, sem As Integer, rango As Integer, loc As Integer) As Double
On Error GoTo Nulo
Dim temp As Double
Dim paque As Double

If rango = 4 Then ' calculo estandar, saca el valordirecto de la db
    temp = FRConsumoMaximoNoPorcionxLocal(ingret, sem - 1, loc)
    paque = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & CotiPrincipalDeIngrediente(ingret))
    FRRequerimientoMaximoxDespachoNoPorcionesxLocalCalculado = redondear_mas(temp / paque) * paque
Else
    FRRequerimientoMaximoxDespachoNoPorcionesxLocalCalculado = FRRequerimientoMaximoxDespachoNoPorcionesxLocal(ingret, sem, rango, loc)
End If

Exit Function

Nulo:
FRRequerimientoMaximoxDespachoNoPorcionesxLocalCalculado = 0

End Function

Function FRConsumoPorcion(coti As Integer, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodCotizacion, Consumo, Semana, FechaPublicada" & _
" FROM CalConsumoPorcionSemanal" & _
" WHERE (((CodCotizacion) = " & coti & ") And ((semana) = " & sem & "))" & _
" ORDER BY FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoPorcion = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoPorcion = 0

End Function

Function FRConsumoPorcionSucursalArchivoAdjunto(coti As Integer, sem As Integer, cola As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodCotizacion, Consumo, Semana, FechaPublicada" & _
" FROM CalConsumoPorcionSemanal" & cola & _
" WHERE (((CodCotizacion) = " & coti & ") And ((semana) = " & sem & "))" & _
" ORDER BY FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoPorcionSucursalArchivoAdjunto = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoPorcionSucursalArchivoAdjunto = 0

End Function

Function FRConsumoPorcionSumaLocales(coti As Integer, sem As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE ((StatusSucursales.Sucursal<>0) AND ((StatusSucursales.Activo)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoPorcionSumaLocales = 0
Do While Not rst.EOF
    FRConsumoPorcionSumaLocales = FRConsumoPorcionSumaLocales + FRConsumoPorcionSucursalArchivoAdjunto(coti, sem, rst("CodLocal"))
    rst.MoveNext
Loop
rst.Close
Exit Function

Nulo:
FRConsumoPorcionSumaLocales = 0
End Function

Function FRConsumoMaximoPorcion(coti As Integer, sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoPorcionSemanal.CodCotizacion, CalConsumoPorcionSemanal.ConsumoMaximo, CalConsumoPorcionSemanal.Semana," & _
" CalConsumoPorcionSemanal.FechaPublicada" & _
" FROM CalConsumoPorcionSemanal" & _
" WHERE (((CalConsumoPorcionSemanal.CodCotizacion) = " & coti & ") And ((CalConsumoPorcionSemanal.semana) = " & sem & "))" & _
" ORDER BY CalConsumoPorcionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoMaximoPorcion = rst("ConsumoMaximo")
rst.Close
     
Exit Function

Nulo:
FRConsumoMaximoPorcion = 0

End Function

Function FRConsumoMaximoPorcionxLocal(coti As Integer, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoPorcionSemanal.CodCotizacion, CalConsumoPorcionSemanal.ConsumoMaximo," & _
" CalConsumoPorcionSemanal.Semana, CalConsumoPorcionSemanal.FechaPublicada, CalConsumoPorcionSemanal.local" & _
" FROM CalConsumoPorcionSemanal IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (CalConsumoPorcionSemanal.CodCotizacion = " & coti & " And CalConsumoPorcionSemanal.semana = " & sem & "" & _
" And CalConsumoPorcionSemanal.local = " & loc & ")" & _
" ORDER BY CalConsumoPorcionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoMaximoPorcionxLocal = rst("ConsumoMaximo")
rst.Close
     
Exit Function

Nulo:
FRConsumoMaximoPorcionxLocal = 0

End Function

Function FRConsumoPorcionxLocal(coti As Integer, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalConsumoPorcionSemanal.CodCotizacion, CalConsumoPorcionSemanal.Consumo," & _
" CalConsumoPorcionSemanal.Semana, CalConsumoPorcionSemanal.FechaPublicada, CalConsumoPorcionSemanal.local" & _
" FROM CalConsumoPorcionSemanal IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (CalConsumoPorcionSemanal.CodCotizacion = " & coti & " And CalConsumoPorcionSemanal.semana = " & sem & "" & _
" And CalConsumoPorcionSemanal.local = " & loc & ")" & _
" ORDER BY CalConsumoPorcionSemanal.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoPorcionxLocal = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoPorcionxLocal = 0

End Function

Function FRRequerimientoMaximoPorciones(coti As Integer, sem As Integer, rango As Integer) As Double
On Error GoTo Nulo
Dim cont As Integer
Dim temp As Double

cont = 0
temp = FRConsumoPorcion(coti, sem - 1 - cont)

Do While cont < rango
If FRConsumoPorcion(coti, sem - 1 - cont) > temp Then
    temp = FRConsumoPorcion(coti, sem - 1 - cont)
End If
cont = cont + 1
Loop
FRRequerimientoMaximoPorciones = temp
Exit Function

Nulo:
FRRequerimientoMaximoPorciones = 0

End Function

Function FRRequerimientoMaximoPorcionesxLocal(coti As Integer, sem As Integer, rango As Integer, loc As Integer) As Double
On Error GoTo Nulo
Dim cont As Integer
Dim temp As Double
Dim temp2 As Double

If rango = 4 Then ' jala el valor directo de la db
    FRRequerimientoMaximoPorcionesxLocal = FRConsumoMaximoPorcionxLocal(coti, sem - 1, loc)

Else
    cont = 0
    temp = FRConsumoPorcionxLocal(coti, sem - 1 - cont, loc)
    
    Do While cont < rango
    temp2 = FRConsumoPorcionxLocal(coti, sem - 1 - cont, loc)
    If temp2 > temp Then
        temp = temp2
    End If
    cont = cont + 1
    Loop
    FRRequerimientoMaximoPorcionesxLocal = temp
    
End If
Exit Function

Nulo:
FRRequerimientoMaximoPorcionesxLocal = 0

End Function

Function FRPedidoSemanaTotalSegunMaximoIngredientePorPorcionesRedondeadoxLocal(ingrex As String, seman As Integer, sucur As Integer, increm As Double, rangox As Integer, cargaseco As Double, cargarefri As Double, cargaconge As Double) As Double
'On Error GoTo Nulo

If ProductoConsumePresentacionPorcion(ingrex) = 0 Then  ' no usa porciones en ninguna receta, solo usa a granel
    FRPedidoSemanaTotalSegunMaximoIngredientePorPorcionesRedondeadoxLocal = 0
    Exit Function
End If

Dim rst As DAO.Recordset
Dim miSQL As String

Dim cantix As Integer
Dim cotix As Long
Dim secoxx As Double
Dim refrix As Double
Dim congex As Double
Dim almac As Integer
Dim paque As Long
Dim requi As Double
Dim converx As Double

secoxx = cargaseco
refrix = cargarefri
congex = cargaconge

miSQL = "SELECT Cotizaciones.CodIngrediente, Cotizaciones.Subproducto, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones WHERE (((Cotizaciones.CodIngrediente)='" & ingrex & "') AND ((Cotizaciones.Subproducto)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantix = rst.RecordCount
rst.MoveFirst

FRPedidoSemanaTotalSegunMaximoIngredientePorPorcionesRedondeadoxLocal = 0

For I = 1 To cantix
    cotix = rst("CodCotizacion")
    paque = IIf(IsNull(DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)), 1, DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotix))
    almac = DLookup("[CodAlmacenamiento]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
    converx = ConversionEstandar(cotix)
    requi = (redondear_mas(absolutopositivo((1 + increm) * factorabastecimiento(almac, secoxx, refrix, congex) * redondear_mas(FRRequerimientoMaximoPorcionesxLocal(PorcionGlobalDePorcion(cotix), seman, rangox, sucur)) - StockSinProcesarxLocal(PorcionDeCotizacion(cotix), seman, sucur)) / paque) * paque) * converx
    FRPedidoSemanaTotalSegunMaximoIngredientePorPorcionesRedondeadoxLocal = FRPedidoSemanaTotalSegunMaximoIngredientePorPorcionesRedondeadoxLocal + requi
    rst.MoveNext
Next I

rst.Close
Exit Function

Nulo:
FRPedidoSemanaTotalSegunMaximoIngredientePorPorcionesRedondeadoxLocal = 0

End Function

Function FRPedidoSemanaTotalSegunMaximoIngredientePorNoPorcionesRedondeadoxLocal(ingrex As String, seman As Integer, sucur As Integer, increm As Double, rangox As Integer, cargaseco As Double, cargarefri As Double, cargaconge As Double) As Double
On Error GoTo Nulo

If ProductoConsumePresentacionNoPorcion(ingrex) = 0 Then
    FRPedidoSemanaTotalSegunMaximoIngredientePorNoPorcionesRedondeadoxLocal = 0

    Exit Function
End If

Dim cantix As Integer
Dim cotix As Long
Dim secoxx As Double
Dim refrix As Double
Dim congex As Double
Dim almac As Integer
Dim paque As Long
Dim requi As Double
Dim converx As Double
Dim requer As Double
Dim stocki As Double
Dim cargax As Double
Dim total As Double

secoxx = cargaseco
refrix = cargarefri
congex = cargaconge


cotix = CotiPrincipalDeIngrediente(ingrex)
paque = IIf(IsNull(DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)), 1, DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotix))
almac = IIf(IsNull(DLookup("[CodAlmacenamiento]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)), 3, DLookup("[CodAlmacenamiento]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)) ' si no existe se considera seco
converx = ConversionEstandar(cotix)
consum = redondear_mas(FRRequerimientoMaximoNoPorcionesxLocal(ingrex, seman, rangox, sucur))
stocki = StockCotizacionlocalconversionestandar(ingrex, seman, sucur) - StockFinalSoloPorcioneslocal(ingrex, seman - 1, sucur)
cargax = factorabastecimiento(almac, secoxx, refrix, congex)
total = (1 + increm) * cargax * consum - stocki

requi = (redondear_mas(absolutopositivo(total) / (converx * paque)) * (converx * paque))
FRPedidoSemanaTotalSegunMaximoIngredientePorNoPorcionesRedondeadoxLocal = requi


Exit Function

Nulo:
FRPedidoSemanaTotalSegunMaximoIngredientePorNoPorcionesRedondeadoxLocal = 0

End Function


Function FRConsumoIngredientexLocal(ING As String, sem As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

Dim consumoTabla As String
consumoTabla = "CalConsumoIngSemanal" & loc

'miSQL = "SELECT CalConsumoIngSemanal.CodIngrediente, CalConsumoIngSemanal.Consumo, CalConsumoIngSemanal.Semana, CalConsumoIngSemanal.FechaPublicada, CalConsumoIngSemanal.local" & _
" FROM CalConsumoIngSemanal IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (CalConsumoIngSemanal.CodIngrediente ='" & ING & "' And CalConsumoIngSemanal.semana = " & sem & " And CalConsumoIngSemanal.local = " & loc & ")" & _
" ORDER BY CalConsumoIngSemanal.FechaPublicada DESC"

miSQL = "SELECT " & consumoTabla & ".CodIngrediente, " & consumoTabla & ".Consumo, " & consumoTabla & ".Semana, " & consumoTabla & ".FechaPublicada" & _
" FROM " & consumoTabla & " IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (" & consumoTabla & ".CodIngrediente ='" & ING & "' And " & consumoTabla & ".semana = " & sem & ")" & _
" ORDER BY " & consumoTabla & ".FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRConsumoIngredientexLocal = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
FRConsumoIngredientexLocal =  0



End Function

Function FRConsumoSemanalMaximoIngredienteLocal(ingred As String, sem As Integer, rango As Integer, loc As Integer) As Double
Dim tempok As Double
Dim maxi As Double
maxi = 0

For I = 1 To rango
    tempok = FRConsumoIngredientexLocal(ingred, sem - I, loc)
    If tempok > maxi Then
        maxi = tempok
    End If
Next I
FRConsumoSemanalMaximoIngredienteLocal = maxi

Exit Function

Nulo:
FRConsumoSemanalMaximoIngredienteLocal = 0

End Function



Function FRPosicionPareto(POS As Integer, sem As Integer) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalParetoSemanal.Posicion, CalParetoSemanal.Semana, CalParetoSemanal.FechaPublicada, CalParetoSemanal.CodIngrediente" & _
" FROM CalParetoSemanal" & _
" WHERE (((CalParetoSemanal.Posicion) = " & POS & ") And ((CalParetoSemanal.semana) = " & sem & "))" & _
" ORDER BY CalParetoSemanal.FechaPublicada DESC"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRPosicionPareto = rst("CodIngrediente")
rst.Close
     
Exit Function

Nulo:
FRPosicionPareto = ""

End Function

Function FRVariacion(sem As Integer) As Double
'faltantes por variacion de stock en costo
On Error Resume Next
Dim rst As DAO.Recordset
Dim miSQL As String

Dim ING As String
Dim SFT, SFR, VAR As Double
Dim Cant As Integer

miSQL = "SELECT [tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10 AS Posicion, FRPosicionPareto([Posicion]," & sem & ") AS CodIngrediente" & _
" FROM tabla_0_al_9, tabla_0_al_9 AS tabla_0_al_9_1" & _
" WHERE ((([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)>0 And ([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)<=25))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

FRVariacion = 0

For I = 1 To Cant
    ING = rst("CodIngrediente")
    SFT = StockIngrediente(ING, sem) + StockCotizacion(ING, sem) + IngresoTotalIngrediente(ING, sem) - (MermaIngrediente(ING, sem) + MermaCotizacion(ING, sem)) - ConsumoSemanalProducto(ING, sem)
    SFR = StockIngrediente(ING, sem + 1) + StockCotizacion(ING, sem + 1)
    VAR = IIf(SFR < SFT, SFT - SFR, 0) * FRCostoUnitarioIngrediente(ING, sem)
    FRVariacion = FRVariacion + VAR
    rst.MoveNext
Next I

rst.Close
Exit Function

Nulo:
FRVariacion = 0

End Function

Function FRVariacionTotal(sem As Integer) As Double
'faltantes por variacion de stock en costo de todos los?ingredientes Variables
On Error Resume Next
Dim rst As DAO.Recordset
Dim miSQL As String

Dim ING As String
Dim SFT, SFR, VAR As Double
Dim Cant As Integer

miSQL = "SELECT DBIngredientes.CodIngrediente, DBIngredientes.TIPO1" & _
" FROM DBIngredientes" & _
" WHERE (((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

FRVariacionTotal = 0

For I = 1 To Cant
    ING = rst("CodIngrediente")
    SFT = StockIngrediente(ING, sem) + StockCotizacion(ING, sem) + IngresoTotalIngrediente(ING, sem) - (MermaIngrediente(ING, sem) + MermaCotizacion(ING, sem)) - ConsumoSemanalProducto(ING, sem)
    SFR = StockIngrediente(ING, sem + 1) + StockCotizacion(ING, sem + 1)
    VAR = (SFR - SFT) * FRCostoUnitarioIngrediente(ING, sem)
    FRVariacionTotal = FRVariacionTotal + VAR
    rst.MoveNext
Next I

rst.Close
Exit Function

Nulo:
FRVariacionTotal = 0

End Function

Function TotalTardanzas(sem As Integer) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([RegistroHorario]![Fecha]) AS semana, Minute([RegistroHorario]![Ingreso])<15 And Minute([RegistroHorario]![Ingreso])>0 AS minutos, Sum(1) AS Total" & _
" FROM RegistroHorario" & _
" GROUP BY numerosemana([RegistroHorario]![Fecha]), Minute([RegistroHorario]![Ingreso])<15 And Minute([RegistroHorario]![Ingreso])>0" & _
" HAVING (((numerosemana([RegistroHorario]![Fecha]))=" & sem & ") AND ((Minute([RegistroHorario]![Ingreso])<15 And Minute([RegistroHorario]![Ingreso])>0)<>0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TotalTardanzas = rst("Total")
rst.Close
     
Exit Function

Nulo:
TotalTardanzas = 0

End Function

Function TotalIngredientes(bat As String) As Double
'Cantidad de ngredientes en un batido

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Sum(SubReceta.Cantidad) AS SumaDeCantidad, SubReceta.CodBatido" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY SubReceta.CodBatido" & _
" HAVING (((SubReceta.CodBatido)='" & bat & "'))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TotalIngredientes = rst("SumaDeCantidad")
rst.Close
     
Exit Function

Nulo:
TotalIngredientes = 0

End Function

Function ProductoCompraVenta(ING As String) As Long
'Condicion si es producto de compra venta o no
' Resultado :   1: Producto Cmpra venta
'               0: No esproducto de compra venta
'Se puede comprar una cotizacion que tenga varios items de ingrediente pero si almenos 1 batido cuenta con 1 solo ingrediente
'se considera como compra venta (suma de ingredientes del batido = 1)
Select Case ING
    Case "O008", "O042", "S018", "S016", "S019", "S024", "P012", "P013", "P014", "S026" 'Productos Pitaya
        ProductoCompraVenta = 1
        Exit Function
    Case "V011", "E004", "E007", "O002", "O006" 'Exceppciones
        ProductoCompraVenta = 0
        Exit Function
End Select
    
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TotalIngredientes([SubReceta]![CodBatido]) AS cond, SubReceta.CodIngrediente" & _
" FROM SubReceta" & _
" GROUP BY TotalIngredientes([SubReceta]![CodBatido]), SubReceta.CodIngrediente" & _
" HAVING (((TotalIngredientes([SubReceta]![CodBatido]))=1) AND ((SubReceta.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ProductoCompraVenta = rst("cond")
rst.Close
     
Exit Function

Nulo:
ProductoCompraVenta = 0

End Function

Function ProductosPromocionSemana(promo As Integer, sem As Integer) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPromocion, numerosemana([NotaDePedido]![Fecha]) AS semana, NotaDePedido.Anulado, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY SubPedido.CodPromocion, numerosemana([NotaDePedido]![Fecha]), NotaDePedido.Anulado" & _
" HAVING (((SubPedido.CodPromocion)=" & promo & ") AND ((numerosemana([NotaDePedido]![Fecha]))=" & sem & ") AND ((NotaDePedido.Anulado)=0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ProductosPromocionSemana = rst("SumaDeCantidad")
rst.Close
     
Exit Function

Nulo:
ProductosPromocionSemana = 0

End Function

Function ProductosPromocionMes(promo As Integer, Mes As Integer, ano As Integer) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPromocion, Month([NotaDePedido]![Fecha]) AS mesi, Year([NotaDePedido]![Fecha]) AS anoi, NotaDePedido.Anulado, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY SubPedido.CodPromocion, Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha]), NotaDePedido.Anulado" & _
" HAVING (SubPedido.CodPromocion=" & promo & " AND Month([NotaDePedido]![Fecha])=" & Mes & " AND Year([NotaDePedido]![Fecha])=" & ano & " AND NotaDePedido.Anulado=0)"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ProductosPromocionMes = rst("SumaDeCantidad")
rst.Close
     
Exit Function

Nulo:
ProductosPromocionMes = 0

End Function

Function ValorizacionMermasSemana(sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT ([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)>0 And ([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)<=25 AS Posicion, Sum((MermaCotizacion(FRPosicionPareto(([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)," & sem & ")," & sem & ")+MermaIngrediente(FRPosicionPareto(([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)," & sem & ")," & sem & "))*FRCostoUnitarioIngrediente(FRPosicionPareto(([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)," & sem & ")," & sem & ")) AS Merma" & _
" FROM tabla_0_al_9, tabla_0_al_9 AS tabla_0_al_9_1" & _
" Group BY([tabla_0_al_9].[N] + [tabla_0_al_9_1].[N] * 10) > 0 And ([tabla_0_al_9].[N] + [tabla_0_al_9_1].[N] * 10) <= 25" & _
" HAVING (((([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)>0 And ([tabla_0_al_9].[N]+[tabla_0_al_9_1].[N]*10)<=25)<>0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ValorizacionMermasSemana = rst("Merma")
rst.Close
     
Exit Function

Nulo:
ValorizacionMermasSemana = 0

End Function

Function ValorizacionMermasSemanaTotal(sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

Dim ING As String
Dim Cant As Integer
Dim Valor As Double

miSQL = "SELECT DBIngredientes.CodIngrediente, DBIngredientes.TIPO1" & _
" FROM DBIngredientes" & _
" WHERE (((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

ValorizacionMermasSemanaTotal = 0

For I = 1 To Cant
    ING = rst("CodIngrediente")
    Valor = (MermaCotizacion(ING, sem) + MermaIngrediente(ING, sem)) * FRCostoUnitarioIngrediente(ING, sem)
    ValorizacionMermasSemanaTotal = ValorizacionMermasSemanaTotal + Valor
    rst.MoveNext
Next I

rst.Close
Exit Function

Nulo:
ValorizacionMermasSemanaTotal = 0

End Function

Function MembresiaVentaSinPromocion(sem As Integer) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Sum(SubPedido.Cantidad) AS SumaDeCantidad, numerosemana([NotaDePedido]![Fecha]) AS semana, SubPedido.CodPromocion, SubPedido.CodBatido, NotaDePedido.Anulado" & _
" FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), SubPedido.CodPromocion, SubPedido.CodBatido, NotaDePedido.Anulado" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sem & ") AND ((SubPedido.CodPromocion)=5) AND ((SubPedido.CodBatido)='TCP') AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MembresiaVentaSinPromocion = rst("SumaDeCantidad")
rst.Close
     
Exit Function

Nulo:
MembresiaVentaSinPromocion = 0

End Function

Function FRCostoBatido(sem As Integer, bat As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodBatido," & _
" Sum([SubReceta]![Cantidad]*" & _
" FactorDeUso([SubReceta]![CodIngrediente],Pedidoplasticovidrio(0))*" & _
" FRCostoUnitarioIngrediente([SubReceta]![CodIngrediente]," & sem & ")) AS Costo" & _
" FROM SubReceta" & _
" GROUP BY SubReceta.CodBatido" & _
" HAVING (((SubReceta.CodBatido)='" & bat & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRCostoBatido = rst("Costo")
rst.Close
     
Exit Function

Nulo:
FRCostoBatido = 0

End Function

Function CosteoPedidosDevueltos(sem As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS semana," & _
" [SubPedido]![CodPromocion]=39 Or [SubPedido]![CodPromocion]=43 AS Devolucion," & _
" NotaDePedido.Anulado, Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])*" & _
" FRCostoUnitarioIngrediente([SubReceta]![CodIngrediente]," & sem & ")) AS Monto" & _
" FROM ((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), [SubPedido]![CodPromocion]=39 Or [SubPedido]![CodPromocion]=43, NotaDePedido.Anulado" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sem & ")" & _
" AND (([SubPedido]![CodPromocion]=39 Or [SubPedido]![CodPromocion]=43)<>0) AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CosteoPedidosDevueltos = rst("Monto")
rst.Close
     
Exit Function

Nulo:
CosteoPedidosDevueltos = 0

End Function
Function VentaxGrupoDiaTipoDeliv(adia As Date, gru As Long, atipo As Integer, deli As Integer, ByRef mont As Double, ByRef acant As Double, montocant As Integer) As Double
'adia: fecha
'gru: grupo de producto venta
'atipo: -1 para pos o tarjeta y 0 para efectivo
'deli: tipo d delivery
'montocant : 1 para monto, 2 para cantidad
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
miSQL = "SELECT NotaDePedido.Anulado, DBBatidos.CodGrupo, NotaDePedido.POS, NotaDePedido.Delivery," & _
" Sum(PrecioReal([SubPedido]![CodSubPedido])*[SubPedido]![Cantidad]*(1+[NotaDepedido]![Propina])) AS Monto, Sum([SubPedido]![Cantidad]) AS Cantidad" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, DBBatidos.CodGrupo, NotaDePedido.POS, NotaDePedido.Delivery, NotaDePedido.Fecha" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((DBBatidos.CodGrupo)=" & gru & ") AND ((NotaDePedido.POS)=" & atipo & ")" & _
" AND ((NotaDePedido.Delivery)=" & deli & ") AND ((NotaDePedido.Fecha)=#" & adia & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
mont = rst("Monto")
acant = rst("Cantidad")
If montocant = 2 Then VentaxGrupoDiaTipoDeliv = acant Else VentaxGrupoDiaTipoDeliv = mont
rst.Close
           
Exit Function

Nulo:
mont = 0
acant = 0
VentaxGrupoDiaTipoDeliv = 0
End Function

Function FRVentaxGrupoDiaTipoDelivInterno(adia As Date, gru As Long, atipo As Integer, adeli As Integer, ByRef mont As Double, ByRef acant As Double, montocant As Integer) As Double  'Interno
'montocant : 1 para monto, 2 para cantidad
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Fecha, CodGrupo, TipoPago, CodDeliv, Cantidad, Monto, FechaPublicada" & _
" FROM CalVentasDiariasxGrupoxTipoxDeli" & _
" WHERE Fecha = #" & adia & "# And CodGrupo = " & gru & " And TipoPago = " & atipo & " And CodDeliv = " & adeli & _
" ORDER BY FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
mont = rst("Monto")
acant = rst("Cantidad")
If montocant = 2 Then FRVentaxGrupoDiaTipoDelivInterno = acant Else FRVentaxGrupoDiaTipoDelivInterno = mont
rst.Close
     
Exit Function

Nulo:
mont = 0
acant = 0
FRVentaxGrupoDiaTipoDelivInterno = 0

End Function

Function FRVentaxGrupoDiaTipoDeliv(adia As Date, gru As Long, atipo As Integer, adeli As Integer, ByRef mont As Double, ByRef acant As Double, montocant As Integer, loc As Integer) As Double  'Global
'montocant : 1 para monto, 2 para cantidad
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Fecha, CodGrupo, TipoPago, CodDeliv, Cantidad, Monto, FechaPublicada, local" & _
" FROM CalVentasDiariasxGrupoxTipoxDeli IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE Fecha = #" & adia & "# And CodGrupo = " & gru & " And TipoPago = " & atipo & " And CodDeliv = " & adeli & " And local = " & loc & _
" ORDER BY FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
mont = rst("Monto")
acant = rst("Cantidad")
If montocant = 2 Then FRVentaxGrupoDiaTipoDeliv = acant Else FRVentaxGrupoDiaTipoDeliv = mont
rst.Close
     
Exit Function

Nulo:
mont = 0
acant = 0
FRVentaxGrupoDiaTipoDeliv = 0

End Function
Function FRVentaxDiaTipoDeliv(adia As Date, atipo As Integer, adeli As Integer, ByRef mont As Double, ByRef acant As Double, montocant As Integer, loc As Integer) As Double  'Global todos los grupos
'montocant : 1 para monto, 2 para cantidad
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim tempmonto As Double
Dim tempcant As Double
Dim temp As Double


miSQL = "SELECT Grupos.CodGrupo FROM Grupos"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

mont = 0
acant = 0
FRVentaxDiaTipoDeliv = 0

For I = 1 To Cant
    temp = FRVentaxGrupoDiaTipoDeliv(adia, rst("CodGrupo"), atipo, adeli, tempmonto, tempcant, 1, loc)
    mont = mont + tempmonto
    acant = acant + tempcant
    rst.MoveNext
Next I
rst.Close

If montocant = 2 Then FRVentaxDiaTipoDeliv = acant Else FRVentaxDiaTipoDeliv = mont
     
Exit Function

Nulo:
mont = 0
acant = 0
FRVentaxDiaTipoDeliv = 0

End Function

Function VentaxGrupoDia(Dia As Date, gru As Long, Valor As Integer) As Double
'Ventas?s totales de un dia , un grupo especifico
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

Select Case Valor
    Case 1

        miSQL = "SELECT DBBatidos.CodGrupo, NotaDePedido.POS, Sum(PrecioReal([SubPedido]![CodSubPedido])*[SubPedido]![Cantidad]*(1+[NotaDepedido]![Propina])) AS Total, NotaDePedido.Anulado" & _
        " FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
        " GROUP BY DBBatidos.CodGrupo, NotaDePedido.Fecha, NotaDePedido.POS, NotaDePedido.Anulado" & _
        " HAVING (((DBBatidos.CodGrupo)=" & gru & ") AND ((NotaDePedido.Fecha)=#" & Dia & "#) AND ((NotaDePedido.POS)<>0) AND ((NotaDePedido.Anulado)=0))"
        Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
        VentaxGrupoDia = rst("Total")
        rst.Close
        
    Case 2

        miSQL = "SELECT DBBatidos.CodGrupo, NotaDePedido.POS, Sum(PrecioReal([SubPedido]![CodSubPedido])*[SubPedido]![Cantidad]*(1+[NotaDepedido]![Propina])) AS Total, NotaDePedido.Anulado" & _
        " FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
        " GROUP BY DBBatidos.CodGrupo, NotaDePedido.Fecha, NotaDePedido.POS, NotaDePedido.Anulado" & _
        " HAVING (((DBBatidos.CodGrupo)=" & gru & ") AND ((NotaDePedido.Fecha)=#" & Dia & "#) AND ((NotaDePedido.POS)=0) AND ((NotaDePedido.Anulado)=0))"
        Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
        VentaxGrupoDia = rst("Total")
        rst.Close
        
    Case 0
    
        miSQL = "SELECT NotaDePedido.Fecha, DBBatidos.CodGrupo, Sum(SubPedido.Cantidad) AS SumaDeCantidad, NotaDePedido.Anulado" & _
        " FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
        " GROUP BY NotaDePedido.Fecha, DBBatidos.CodGrupo, NotaDePedido.Anulado" & _
        " HAVING (((NotaDePedido.Fecha)=#" & Dia & "#) AND ((DBBatidos.CodGrupo)=" & gru & ") AND ((NotaDePedido.Anulado)=0))"
        Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
        VentaxGrupoDia = rst("SumaDeCantidad")
        rst.Close
    
    Case Else
    
        MsgBox "Codigo valor erroneo"
        
End Select
     
Exit Function

Nulo:
VentaxGrupoDia = 0

End Function

Function FRVentaxGrupoDiaInterno(Dia As Date, gru As Long, Valor As Integer) As Double  'Interno
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalVentasDiariasxGrupo.Fecha, CalVentasDiariasxGrupo.CodGrupo, CalVentasDiariasxGrupo.Cantidad," & _
" CalVentasDiariasxGrupo.MontoNoPOS, CalVentasDiariasxGrupo.MontoPOS, CalVentasDiariasxGrupo.FechaPublicada" & _
" FROM CalVentasDiariasxGrupo" & _
" WHERE (((CalVentasDiariasxGrupo.Fecha) = #" & Dia & "#) And ((CalVentasDiariasxGrupo.CodGrupo) = " & gru & "))" & _
" ORDER BY CalVentasDiariasxGrupo.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Select Case Valor
    Case 0
        FRVentaxGrupoDiaInterno = rst("Cantidad")
    Case 1
        FRVentaxGrupoDiaInterno = rst("MontoPOS")
    Case 2
        FRVentaxGrupoDiaInterno = rst("MontoNoPOS")
    Case 3
        FRVentaxGrupoDiaInterno = rst("MontoNoPOS") + rst("MontoPOS")
    Case Else
        MsgBox "Valor solicitado errado"
End Select

rst.Close
     
Exit Function

Nulo:
FRVentaxGrupoDiaInterno = 0

End Function

Function FRVentaxGrupoDia(Dia As Date, gru As Long, Valor As Integer, loc As Integer) As Double  'Global
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CalVentasDiariasxGrupo.Fecha, CalVentasDiariasxGrupo.CodGrupo, CalVentasDiariasxGrupo.Cantidad, CalVentasDiariasxGrupo.MontoNoPOS, CalVentasDiariasxGrupo.MontoPOS, CalVentasDiariasxGrupo.FechaPublicada, CalVentasDiariasxGrupo.local" & _
" FROM CalVentasDiariasxGrupo IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((CalVentasDiariasxGrupo.Fecha) = #" & Dia & "#) And ((CalVentasDiariasxGrupo.CodGrupo) = " & gru & ") And ((CalVentasDiariasxGrupo.local) = " & loc & "))" & _
" ORDER BY CalVentasDiariasxGrupo.FechaPublicada DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Select Case Valor
    Case 0
        FRVentaxGrupoDia = rst("Cantidad")
    Case 1
        FRVentaxGrupoDia = rst("MontoPOS")
    Case 2
        FRVentaxGrupoDia = rst("MontoNoPOS")
    Case 3
        FRVentaxGrupoDia = rst("MontoNoPOS") + rst("MontoPOS")
    Case Else
        MsgBox "Valor solicitado errado"
End Select

rst.Close
     
Exit Function

Nulo:
FRVentaxGrupoDia = 0

End Function

Function VentaTotalesMontoGuardadoxLocal(sema As Integer, loci As Integer) As Double  'Global

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotadePedido.local, numerosemana([NotaDePedido]![Fecha]) AS Expr1," & _
" Sum(NotadePedido.TotalGuardado) AS SumaDeTotalGuardado" & _
" FROM NotadePedido IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY NotadePedido.local, numerosemana([NotaDePedido]![Fecha])" & _
" HAVING (((NotadePedido.local)=" & loci & ") AND ((numerosemana([NotaDePedido]![Fecha]))=" & sema & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentaTotalesMontoGuardadoxLocal = rst("SumaDeTotalGuardado")
rst.Close
     
Exit Function

Nulo:
VentaTotalesMontoGuardadoxLocal = 0

End Function

Function VentaTotalesMontoGuardadoTodosLocales(sema As Integer) As Double  'Global

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Expr1," & _
" Sum(NotadePedido.TotalGuardado) AS SumaDeTotalGuardado" & _
" FROM NotadePedido IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha])" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sema & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentaTotalesMontoGuardadoTodosLocales = rst("SumaDeTotalGuardado")
rst.Close
     
Exit Function

Nulo:
VentaTotalesMontoGuardadoTodosLocales = 0

End Function

Function FRVentaxGrupoSemana(sem As Integer, gru As Long, Valor As Integer) As Double  'Interno
'Valor=0: Cantidad
'Valor=1: Monto POS
'Valor=2: Monto Sin POS
'Valor=3: Monto Sin POS + POS

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([FechaSistema]![Dates]) AS sema, FechaSistema.Dates" & _
" FROM FechaSistema" & _
" WHERE (((numerosemana([FechaSistema]![Dates]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst

FRVentaxGrupoSemana = 0

For I = 1 To 7 ' 7 dias de la semana
    FRVentaxGrupoSemana = FRVentaxGrupoSemana + FRVentaxGrupoDiaInterno(rst("Dates"), gru, Valor)
    rst.MoveNext
Next I

rst.Close
     
Exit Function

Nulo:
FRVentaxGrupoSemana = 0

End Function

Function FechaDepositoAnterior(fech As Date) As Date

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TOP 1 Depositos.Fecha FROM Depositos" & _
" WHERE (((Depositos.Fecha) < #" & fech & "#)) ORDER BY Depositos.Fecha DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FechaDepositoAnterior = rst("Fecha")
rst.Close
     
Exit Function

Nulo:
FechaDepositoAnterior = 0

End Function

Function PagoInfluencers(sem As Integer) As Double
'Codigos considerados:
'2107:Nancy,  2098:Andrea Ruiz,  2105:Farah,  2086:Niniveth
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [NotaDePedido]![CodCliente]=2107 Or [NotaDePedido]![CodCliente]=2098 Or [NotaDePedido]![CodCliente]=2105 Or [NotaDePedido]![CodCliente]=2086 AS clientes, SubPedido.CodPromocion, numerosemana([NotaDePedido]![Fecha]) AS semana, NotaDePedido.Anulado, Sum([SubPedido]![Cantidad]*[DBBatidos]![Precio]) AS Total" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY [NotaDePedido]![CodCliente]=2107 Or [NotaDePedido]![CodCliente]=2098 Or [NotaDePedido]![CodCliente]=2105 Or [NotaDePedido]![CodCliente]=2086, SubPedido.CodPromocion, numerosemana([NotaDePedido]![Fecha]), NotaDePedido.Anulado" & _
" HAVING ((([NotaDePedido]![CodCliente]=2107 Or [NotaDePedido]![CodCliente]=2098 Or [NotaDePedido]![CodCliente]=2105 Or [NotaDePedido]![CodCliente]=2086)<>0) AND ((SubPedido.CodPromocion)=22) AND ((numerosemana([NotaDePedido]![Fecha]))=" & sem & ") AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PagoInfluencers = rst("Total")
rst.Close
     
Exit Function

Nulo:
PagoInfluencers = 0

End Function

Function FRCostoIFIngrediente(sem As Integer) As Double
'Costo de los items de ingredientes realizados los domingos, de una semana especifica
'COSTOS VARIABLES
'CONSTO INVNTARIO FINAL DE SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular la suma de los costos de ingredientes por cantidad de ingredientes del inventario de domingo
miSQL = "SELECT DBIngredientes.TIPO1, numerosemana([Inventario Ingrediente]![Fecha]) AS semana, Sum([Inventario Ingrediente]![Cantidad]*FRCostoUnitarioIngrediente([Inventario Ingrediente]![CodIngrediente],numerosemana([Inventario Ingrediente]![Fecha]))) AS Costeo" & _
" FROM DBIngredientes INNER JOIN [Inventario Ingrediente] ON DBIngredientes.CodIngrediente = [Inventario Ingrediente].CodIngrediente" & _
" GROUP BY DBIngredientes.TIPO1, numerosemana([Inventario Ingrediente]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((numerosemana([Inventario Ingrediente]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRCostoIFIngrediente = rst("Costeo")
rst.Close

Exit Function

Nulo:
FRCostoIFIngrediente = 0

End Function

Function FRCostoIFCotizacion(sem As Integer) As Double
'Costo de los items de cotizacion realizados los domingos, de una semana especifica
'COSTOS VARIABLES
'CONSTO INVNTARIO FINAL DE SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular la suma de los costos de ingredientes por cantidad de ingredientes del inventario de domingo
miSQL = "SELECT DBIngredientes.TIPO1, numerosemana([Inventario Cotizacion]![Fecha]) AS semana, Sum([Inventario Cotizacion]![Cantidad]*FRValorCotizacion([Inventario Cotizacion]![CodCotizacion],numerosemana([Inventario Cotizacion]![Fecha]))) AS Costeo" & _
" FROM [Inventario Cotizacion] INNER JOIN (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, numerosemana([Inventario Cotizacion]![Fecha])" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRCostoIFCotizacion = rst("Costeo")
rst.Close

Exit Function

Nulo:
FRCostoIFCotizacion = 0

End Function
'=================================Costo de Ventas====================================
Function FRCostoConsumoTeoricoSemana(sem As Integer) As Double 'local
''Suma de costo de cantidades consumidas teoricos*

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS semana, NotaDePedido.Anulado, SubReceta.CodIngrediente," & _
" Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Cantidad" & _
" FROM (SubPedido INNER JOIN (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), NotaDePedido.Anulado, SubReceta.CodIngrediente" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sem & ") AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioIngrediente(rst("CodIngrediente"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoConsumoTeoricoSemana = suma
Exit Function

Nulo:
FRCostoConsumoTeoricoSemana = 0

End Function

Function FRCostoIngresosSemana(sem As Integer) As Double 'local

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT DBIngredientes.TIPO1, numerosemana([IngresosPitaya]![Fecha]) AS semana, IngresosPitaya.CodCotizacion, Sum([IngresosPitaya]![Cantidad]) AS Cantidad" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, numerosemana([IngresosPitaya]![Fecha]), IngresosPitaya.CodCotizacion" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((numerosemana([IngresosPitaya]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioCotizacion(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoIngresosSemana = suma
Exit Function

Nulo:
FRCostoIngresosSemana = 0

End Function

Function FRCostoMermasCotizacionSemana(sem As Integer) As Double 'local
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT DBIngredientes.TIPO1, numerosemana([Merma Cotizacion]![Fecha]) AS semana, [Merma Cotizacion].CodCotizacion, Sum([Merma Cotizacion]![Cantidad]) AS Cantidad" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN [Merma Cotizacion] ON Cotizaciones.CodCotizacion = [Merma Cotizacion].CodCotizacion" & _
" GROUP BY DBIngredientes.TIPO1, numerosemana([Merma Cotizacion]![Fecha]), [Merma Cotizacion].CodCotizacion" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((numerosemana([Merma Cotizacion]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioCotizacion(rst("CodCotizacion"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoMermasCotizacionSemana = suma
Exit Function

Nulo:
FRCostoMermasCotizacionSemana = 0

End Function

Function FRCostoMermasIngredienteSemana(sem As Integer) As Double 'local
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT DBIngredientes.TIPO1, numerosemana([Merma Unidad]![Fecha]) AS semana, [Merma Unidad].CodIngrediente, Sum([Merma Unidad]![Cantidad]) AS Cantidad" & _
" FROM DBIngredientes INNER JOIN [Merma Unidad] ON DBIngredientes.CodIngrediente = [Merma Unidad].CodIngrediente" & _
" GROUP BY DBIngredientes.TIPO1, numerosemana([Merma Unidad]![Fecha]), [Merma Unidad].CodIngrediente" & _
" HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((numerosemana([Merma Unidad]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    suma = suma + rst("Cantidad") * FRCostoUnitarioIngrediente(rst("CodIngrediente"), rst("semana"))
    rst.MoveNext
Loop
rst.Close

FRCostoMermasIngredienteSemana = suma
Exit Function

Nulo:
FRCostoMermasIngredienteSemana = 0

End Function

'====================================================================================================
'=========================FR Costo Ingrediente Funcion Resumida========================


Function FRCostoIngredienteProcesamiento(ING As String, sem As Integer, modo As Integer) As Double
'Costo de procesamiento de un ingredinte en una semana especifica
'COSO DE LA CANTIDAD DE INGREDIENTE PROCESADO EN LA SEMANA Y LA CANTIDAD PROCESADA EN LA SEMANA4
'UNIDADES CON CONVERSION = 0   ' INGRESO POR PROCESAMIENTO

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular los valores de total procesdo de ingrediente (1) y costo total de lo procesado en cotizacion (2)
miSQL = "SELECT numerosemana([Procesamiento]![Fecha]) AS semana, Cotizaciones.CodIngrediente, Sum(Procesamiento.MedidaFinal) AS SumaDeMedidaFinal," & _
" Sum([Procesamiento]![Cantidad]*FRValorCotizacion([Procesamiento]![CodCotizacion],numerosemana([Procesamiento]![Fecha]))) AS CostoTotal" & _
" FROM Procesamiento INNER JOIN Cotizaciones ON Procesamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([Procesamiento]![Fecha]), Cotizaciones.CodIngrediente" & _
" HAVING (((numerosemana([Procesamiento]![Fecha]))=" & sem & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If modo = 1 Then
FRCostoIngredienteProcesamiento = rst("SumaDeMedidaFinal")
rst.Close
Else
FRCostoIngredienteProcesamiento = rst("CostoTotal")
rst.Close
End If


Exit Function

Nulo:
FRCostoIngredienteProcesamiento = 0

End Function

Function FRCostoIngredienteIngreso(ING As String, sem As Integer, modo As Integer) As Double
'Costo de ingresos de un ingredinte en una semana especifica
'COSTO DE LO Q INGRESA EN LA SEMANA Y TOTAL EN VALOR ING DE LO QUE INGRESA

' UNIDADES CON CONVERSION NO 0  ' INGRESO NATURAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular los valores de total ingresado de ingrediente (1) y costo total de lo ingresado en cotizacion (2)
miSQL = "SELECT numerosemana([IngresosPitaya]![Fecha]) AS semana, Cotizaciones.CodIngrediente, [Cotizaciones]![Conversion]=0 AS conversion," & _
" Sum([IngresosPitaya]![Cantidad]*FRConversionCotizacion([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS IngresoTotal," & _
" Sum([IngresosPitaya]![Cantidad]*FRValorCotizacion([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS CostoTotal" & _
" FROM IngresosPitaya" & _
" INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([IngresosPitaya]![Fecha]), Cotizaciones.CodIngrediente, [Cotizaciones]![Conversion]=0" & _
" HAVING (((numerosemana([IngresosPitaya]![Fecha]))=" & sem & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "') AND (([Cotizaciones]![Conversion]=0)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If modo = 1 Then
FRCostoIngredienteIngreso = rst("IngresoTotal")
rst.Close
Else
FRCostoIngredienteIngreso = rst("CostoTotal")
rst.Close
End If


Exit Function

Nulo:
FRCostoIngredienteIngreso = 0

End Function

Function FRCostoIngredienteGlobal(ING As String, sem As Integer) As Double

' Solo cuando ya haya sido guardado conversion y costu cotizacion 
' Costo Unitario de ingrediente acorde a ingresos y procesamientos de una semana especifica
' SI NO HAY DATOS EN LA SEMANA AGARRA DE SEMANA ATERIOR

On Error GoTo Nulo

If sem < 160 Then 'semana de referencia inicio de reportes semanales
FRCostoIngredienteGlobal = CostoIngredienteGlobal(ING, sem)
Exit Function
End If

FRCostoIngredienteGlobal = ((StockCotizacionPorcionable(ING, sem) + StockIngrediente(ING, sem)) * FRCostoUnitarioIngrediente(ING, sem - 1) + FRCostoIngredienteProcesamiento(ING, sem, 2) + FRCostoIngredienteIngreso(ING, sem, 2)) / (FRCostoIngredienteProcesamiento(ING, sem, 1) + FRCostoIngredienteIngreso(ING, sem, 1) + (StockCotizacionPorcionable(ING, sem) + StockIngrediente(ING, sem)))

Exit Function

Nulo:
FRCostoIngredienteGlobal = FRCostoUnitarioIngrediente(ING, sem - 1)

End Function

'======================================================================================
'=============================================================================================

Function tipoproductoreporteprincipal(ingre As String) As Integer 'local
'si es 0, solo tiene porciones
'<>0 , tiene porciones y no porciones

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [DBBatidos]![CodGrupo]=7 AS Expr2, SubReceta.CodIngrediente, DBBatidos.Vigencia, Sum(IIf(IsNull([SubReceta]![codporcion])<>0,1,0)) AS Expr1" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY [DBBatidos]![CodGrupo]=7, SubReceta.CodIngrediente, DBBatidos.Vigencia" & _
" HAVING ((([DBBatidos]![CodGrupo]=7)=0) AND ((SubReceta.CodIngrediente)='" & ingre & "') AND ((DBBatidos.Vigencia)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
tipoproductoreporteprincipal = IIf(rst("Expr1") <> 0, 1, 0)
rst.Close


Exit Function

Nulo:
tipoproductoreporteprincipal = 1

End Function

