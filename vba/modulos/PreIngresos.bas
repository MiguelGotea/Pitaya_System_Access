' ==========================================================
' Modulo  : PreIngresos
' Tipo    : 1  |  Lineas: 703
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database

Sub AutoIngresoDatosPorciones(seman As Integer, rangi As Integer, incres As Double, locs As Integer, despa As Integer, preis As Long, preic As Long, contadorerror As Integer)

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer
Dim ind As Integer
Dim prei As Long

Dim sema As Integer
Dim rang As Integer
Dim loca As Integer
Dim pedi As Double

Dim redreq As Integer
Dim cm As Double
Dim SI As Double
Dim pendi As Integer
Dim abas As Integer

Dim datsec1 As Double
Dim datref1 As Double
Dim datcon1 As Double
Dim datsec2 As Double
Dim datref2 As Double
Dim datcon2 As Double

sema = seman
rang = rangi
loca = locs

datsec1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datref1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=1 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datcon1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")

datsec2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datref2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=1 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datcon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")

miSQL = "SELECT TiposVariables.Orden, DBIngredientes.Nombre, TiposVariables.Control, DBBatidos.Vigencia," & _
" SubReceta.codporcion, [DBBatidos]![CodGrupo]=7 AS nomostrador, PorcionDentroDeMezcla([SubReceta]![codporcion]) AS mezcla," & _
" DBIngredientes.Tipo, DLookUp('[CodAlmacenamiento]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) AS almacen," & _
" DLookUp('[PaquetePorciones]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) AS paqueteporciones" & _
" FROM ((SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente)" & _
" INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" GROUP BY TiposVariables.Orden, DBIngredientes.Nombre, TiposVariables.Control, DBBatidos.Vigencia, SubReceta.codporcion," & _
" [DBBatidos]![CodGrupo]=7, PorcionDentroDeMezcla([SubReceta]![codporcion]), DBIngredientes.Tipo," & _
" DLookUp('[CodAlmacenamiento]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])," & _
" DLookUp('[PaquetePorciones]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])" & _
" HAVING (((TiposVariables.Control) = True) And ((DBBatidos.Vigencia) = True) And ((SubReceta.codporcion) Is Not Null)" & _
" And (([DBBatidos]![CodGrupo] = 7) = False) And ((PorcionDentroDeMezcla([SubReceta]![codporcion])) = 0))" & _
" ORDER BY TiposVariables.Orden, DBIngredientes.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For ind = 1 To canti
    If ind >= contadorerror Then
        abas = rst("almacen")
        pedi = factorabastecimiento(abas, datsec1, datref1, datcon1)
        cm = redondear_mas(FRRequerimientoMaximoPorcionesxLocal(PorcionGlobalDePorcion(rst("codporcion")), sema, rang, loca))
        SI = StockSinProcesarxLocal(PorcionDeCotizacion(rst("codporcion")), sema, loca)
        redreq = redondear_mas(absolutopositivo((1 + incres) * pedi * cm - SI) / rst("PaquetePorciones")) * rst("PaquetePorciones")
        
        If despa = 2 Then
            pedi = factorabastecimiento(abas, datsec2, datref2, datcon2)
            redreq = redondear_mas(absolutopositivo((1 + incres) * pedi * cm - SI) / rst("PaquetePorciones")) * rst("PaquetePorciones") - redreq
        End If
        
        Select Case abas
            Case 1, 2 'congelados
                prei = preic
            Case 3 'secos refrigerados
                prei = preis
            Case Else
                prei = preis
        End Select

        If redreq > 0 Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & rst("codporcion") & ", " & redreq & ", " & prei & ")"
            DoCmd.SetWarnings True
            
            If ExisteMezcla(rst("codporcion")) = 1 Then
                ' Agregar la mezcla original
                ' Obtener CodCotizacionPorcion relacionados
                Dim rsPorciones As DAO.Recordset
                Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & rst("codporcion"))
        
                ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
                Do While Not rsPorciones.EOF
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
                        " values (" & rsPorciones!CodCotizacionPorcion & ", " & redreq & ", " & prei & ")"
                    rsPorciones.MoveNext
                    DoCmd.SetWarnings True
                Loop
        
                rsPorciones.Close
            End If

        End If
    End If
    rst.MoveNext
Next ind

rst.Close
Exit Sub

AgainAgain:
rst.Close
Call AutoIngresoDatosPorciones(seman, rangi, incres, locs, despa, preis, preic, ind)
End Sub

Sub AutoIngresoDatosNoPerecibles(seman As Integer, rangi As Integer, incres As Double, locs As Integer, despa As Integer, preis As Long, preic As Long, contadorerror As Integer)
'NO PORCIONES
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer
Dim ind As Integer

Dim sema As Integer
Dim rang As Integer
Dim loca As Integer
Dim pedi As Double
Dim cantipaq As Integer
Dim conver As Double

Dim redreq As Integer
Dim cm As Double
Dim SI As Double
Dim pendi As Integer
Dim abas As Integer

Dim datsec1 As Double
Dim datref1 As Double
Dim datcon1 As Double
Dim datsec2 As Double
Dim datref2 As Double
Dim datcon2 As Double

sema = seman
rang = rangi
loca = locs

datsec1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=2 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datref1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=2 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datcon1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=2 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")

datsec2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=2 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datref2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=2 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datcon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=2 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")

miSQL = "SELECT TiposVariables.Orden, DBIngredientes.Nombre, TiposVariables.Control, DBBatidos.Vigencia," & _
" SubReceta.codporcion, [DBBatidos]![CodGrupo]=7 AS nomostrador, DBIngredientes.compralocal, DBIngredientes.Tipo," & _
" Cotizaciones.CodAlmacenamiento, SubReceta.CodIngrediente, CotiPrincipalDeIngrediente([SubReceta]![CodIngrediente]) AS coti" & _
" FROM (((DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente)" & _
" INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" GROUP BY TiposVariables.Orden, DBIngredientes.Nombre, TiposVariables.Control," & _
" DBBatidos.Vigencia, SubReceta.codporcion, [DBBatidos]![CodGrupo]=7, DBIngredientes.compralocal," & _
" DBIngredientes.Tipo, Cotizaciones.CodAlmacenamiento, SubReceta.CodIngrediente," & _
" CotiPrincipalDeIngrediente([SubReceta]![CodIngrediente])" & _
" HAVING (((TiposVariables.Control) = True) And ((DBBatidos.Vigencia) = True)" & _
" And ((SubReceta.codporcion) Is Null) And (([DBBatidos]![CodGrupo] = 7) = False) And ((DBIngredientes.compralocal) = False))" & _
" ORDER BY TiposVariables.Orden, DBIngredientes.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For ind = 1 To canti
    If ind >= contadorerror Then
        abas = rst("CodAlmacenamiento")
        pedi = factorabastecimiento(abas, datsec1, datref1, datcon1)
        cantipaq = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & rst("coti"))
        conver = conversioncalculado(rst("coti"), sema - 1)
        
        cm = FRRequerimientoMaximoNoPorcionesxLocal(rst("CodIngrediente"), sema, rang, loca)
        SI = StockCotizacionlocalconversionestandar(rst("CodIngrediente"), sema, loca) - StockFinalSoloPorcioneslocal(rst("CodIngrediente"), sema - 1, loca)
        
        redreq = redondear_mas(absolutopositivo((cm * pedi * (1 + incres) - SI) / conver) / cantipaq) * cantipaq

        If despa = 2 Then
            pedi = factorabastecimiento(abas, datsec2, datref2, datcon2)
            redreq = redondear_mas(absolutopositivo((cm * pedi * (1 + incres) - SI - redreq * conver) / conver) / cantipaq) * cantipaq
        End If
        
        Select Case abas
            Case 1
                prei = preic
            Case 2, 3
                prei = preis
            Case Else
                prei = preis
        End Select
        
        If redreq > 0 Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & rst("coti") & ", " & redreq & ", " & prei & ")"
            DoCmd.SetWarnings True

        End If
    End If
    rst.MoveNext
Next ind

rst.Close
Exit Sub

AgainAgain:
rst.Close
Call AutoIngresoDatosNoPerecibles(seman, rangi, incres, locs, despa, preis, preic, ind)
End Sub

Sub AutoIngresoDatosMostrador(seman As Integer, rangi As Integer, incres As Double, locs As Integer, despa As Integer, preis As Long, preic As Long, contadorerror As Integer)

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer
Dim ind As Integer

Dim sema As Integer
Dim rang As Integer
Dim loca As Integer
Dim pedi As Double

Dim redreq As Integer
Dim cm As Double
Dim SI As Double
Dim pendi As Integer
Dim abas As Integer
Dim cotis As Integer
Dim cantispaq As Integer

Dim datsec1 As Double
Dim datref1 As Double
Dim datcon1 As Double
Dim datsec2 As Double
Dim datref2 As Double
Dim datcon2 As Double

sema = seman
rang = rangi
loca = locs

datsec1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datref1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datcon1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(sema) & "'")

datsec2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datref2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")
datcon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(sema) & "'")

'miSQL = "SELECT DBBatidos.Marca," & _
'" DBBatidos.CodBatido," & _
'" DBBatidos.Nombre," & _
'" DBBatidos.Medida," & _
'" DBBatidos.Vigencia," & _
'" DBBatidos.CodGrupo" & _
'" FROM DBBatidos WHERE (((DBBatidos.CodBatido) Not Like '*d')" & _
'" AND ((DBBatidos.Vigencia)<>0) AND ((DBBatidos.CodGrupo)=7))" & _
'" ORDER BY DBBatidos.Marca, DBBatidos.CodBatido, DBBatidos.Nombre"

miSQL = "SELECT DBBatidos.Marca," & _
" DBBatidos.Nombre," & _
" DBBatidos.Medida," & _
" DBBatidos.Vigencia," & _
" DBBatidos.CodGrupo" & _
" FROM DBBatidos" & _
" GROUP BY DBBatidos.Marca, DBBatidos.Nombre," & _
" DBBatidos.Medida, DBBatidos.Vigencia, DBBatidos.CodGrupo" & _
" HAVING (((DBBatidos.Vigencia) = True) And ((DBBatidos.CodGrupo) = 7))" & _
" ORDER BY DBBatidos.Marca, DBBatidos.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For ind = 1 To canti
    If ind >= contadorerror Then
        'cotiinsumoclavexnombreproductoventa
        cotis = cotiinsumoclavexnombreproductoventa(rst("Nombre"))
        abas = DLookup("[CodAlmacenamiento]", "[Cotizaciones]", "[CodCotizacion]=" & cotis)
        pedi = factorabastecimiento(abas, datsec1, datref1, datcon1)
        cantispaq = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotis)
        cm = ventaMaximoproductosemanalocal(rst("Nombre"), rst("Medida"), sema - 1, rang, loca)
        '=StockSinProcesarxLocalxPorductoVenta([Formularios]![ConsumoPitayaStore]![Nombre],[asemana],[alocali])
        SI = StockSinProcesarxLocalxPorductoVenta(rst("Nombre"), sema, loca)
        redreq = redondear_mas(absolutopositivo(cm * (1 + incres) * pedi - SI) / cantispaq) * cantispaq
        If despa = 2 Then
            pedi = factorabastecimiento(abas, datsec2, datref2, datcon2)
            redreq = redondear_mas(absolutopositivo(cm * (1 + incres) * pedi - SI - redreq) / cantispaq) * cantispaq
        End If

        Select Case abas
            Case 1, 2 'congelados
                prei = preis
            Case 3 'secos refrigerados
                prei = preis
            Case Else
                prei = preis
        End Select
        
        If redreq > 0 Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & cotis & ", " & redreq & ", " & prei & ")"
            DoCmd.SetWarnings True
            
            If ExisteMezcla(cotis) = 1 Then
                ' Agregar la mezcla original
                ' Obtener CodCotizacionPorcion relacionados
                Dim rsPorciones As DAO.Recordset
                Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & cotis)
            
                ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
                Do While Not rsPorciones.EOF
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
                        " values (" & rsPorciones!CodCotizacionPorcion & ", " & redreq & ", " & prei & ")"
                    rsPorciones.MoveNext
                    DoCmd.SetWarnings True
                Loop
            
                rsPorciones.Close
            End If

        End If
    End If
    rst.MoveNext
Next ind

rst.Close
Exit Sub

AgainAgain:
rst.Close
Call AutoIngresoDatosMostrador(seman, rangi, incres, locs, despa, preis, preic, ind)
End Sub


Sub AutoIngresoDatosInsumosFijos(seman As Integer, locs As Integer, prei As Long)

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer
Dim ind As Integer

Dim paqpor As Double
Dim redreq As Double
Dim codal As Integer
Dim cm As Double
Dim inve As Double
Dim carga As Double

miSQL = "SELECT DBIngredientes.Tipo, nombreproductocoti([Cotizaciones]![CodCotizacion]) AS Nombre," & _
" Cotizaciones.Prioridad, DBIngredientes.Consumible, Cotizaciones.CodCotizacion," & _
" Cotizaciones.PaquetePorciones, Cotizaciones.CodAlmacenamiento" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (((Cotizaciones.Prioridad) = True) And ((DBIngredientes.Consumible) = True))" & _
" ORDER BY DBIngredientes.Tipo, nombreproductocoti([Cotizaciones]![CodCotizacion])"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For ind = 1 To canti
    paqpor = rst("PaquetePorciones")
    codal = rst("CodAlmacenamiento")
    cm = EstandarInsumoFijoCotizacionxLocal(rst("CodCotizacion"), locs)
    inve = StockSinProcesarxLocal(rst("CodCotizacion"), seman, locs)
    carga = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & locs & " AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(seman) & "'")
    redreq = redondear_mas(absolutopositivo(factorabastecimiento(codal, carga, carga, carga) * cm / 4 - inve) / paqpor) * paqpor
    If redreq > 0 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
        " values (" & rst("CodCotizacion") & ", " & redreq & ", " & prei & ")"
        DoCmd.SetWarnings True
    End If
    rst.MoveNext
Next ind

rst.Close
Exit Sub

AgainAgain:
rst.Close
MsgBox "No existe ningun producto solicitado por la sucursal"
End Sub

Function statuspreingreso(preinge As Long) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

If ContadorItemsPreingreso(preinge) = 0 Then ' si es lista vacia cuenta como ingresado
    statuspreingreso = -1
    Exit Function
End If

miSQL = "SELECT StatusPreingreso.CodPreIngresoPitaya, StatusPreingreso.Status FROM statuspreingreso" & _
" WHERE (((StatusPreingreso.CodPreIngresoPitaya)=" & preinge & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
statuspreingreso = rst("Status")
rst.Close

Exit Function

Nulo:
statuspreingreso = 0

End Function

Function statuscambiospreingreso(preinge As Long) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusCambiosPreingreso.CodCambiosPreIngresoPitaya, StatusCambiosPreingreso.Status FROM StatusCambiosPreingreso" & _
" WHERE (((StatusCambiosPreingreso.CodCambiosPreIngresoPitaya)=" & preinge & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
statuscambiospreingreso = rst("Status")
rst.Close

Exit Function

Nulo:
statuscambiospreingreso = 0

End Function

Function existepreingresopendientefecha(fechi As Date) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT PreIngresoPitaya.Destino, PreIngresoPitaya.Fecha, statuspreingreso([PreIngresoPitaya]![CodPreIngresoPitaya]) AS status," & _
" PreIngresoPitaya.Validado, Sum(1) AS cont" & _
" FROM PreIngresoPitaya" & _
" GROUP BY PreIngresoPitaya.Destino, PreIngresoPitaya.Fecha," & _
" statuspreingreso([PreIngresoPitaya]![CodPreIngresoPitaya]), PreIngresoPitaya.Validado" & _
" HAVING (((PreIngresoPitaya.Destino)='Pitaya ' & codigolocal())" & _
" AND ((PreIngresoPitaya.Fecha)=#" & fechi & "#)" & _
" AND ((statuspreingreso([PreIngresoPitaya]![CodPreIngresoPitaya]))=0)" & _
" AND ((PreIngresoPitaya.Validado)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
existepreingresopendientefecha = rst("cont")
rst.Close

Exit Function

Nulo:
existepreingresopendientefecha = 0
' todos los preingresos estan ingresados
End Function

Function existecambiospreingresopendiente(preinge As Long) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CambiosPreIngresosPitaya.CodPreIngresoPitaya," & _
" statuscambiospreingreso([CambiosPreIngresosPitaya]![CodCambiosPreIngresoPitaya]) AS Expr1, 1 AS cont" & _
" FROM CambiosPreIngresosPitaya" & _
" WHERE (((CambiosPreIngresosPitaya.CodPreIngresoPitaya)=" & preinge & ")" & _
" AND ((statuscambiospreingreso([CambiosPreIngresosPitaya]![CodCambiosPreIngresoPitaya]))=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
existecambiospreingresopendiente = rst("cont")
rst.Close

Exit Function

Nulo:
existecambiospreingresopendiente = 0
' todos los preingresos estan ingresados
End Function

Function crearpreingreso(fechas As Date, sucu As Integer) As Long
On Error GoTo Nulo

'Nuevo preingreso
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO PreIngresoPitaya(Fecha, Hora, Destino)" & _
" values (#" & fechas & "#, #" & Time & "#,'Pitaya " & sucu & "')"
DoCmd.SetWarnings True

Dim codpre As Long
codpre = ultimocodigopreingreso()
crearpreingreso = codpre
Exit Function
Nulo:
MsgBox "No se ha creado registro, vuelva a intentarlo"
crearpreingreso = 0
End Function

Function PedidoInsumoFijoCotizacionSemana(coti As Integer, sema As Integer) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT PedidoInsumosFijosSemana.CodCotizacion, PedidoInsumosFijosSemana.Semana," & _
" Sum(PedidoInsumosFijosSemana.Cantidad) AS SumaDeCantidad" & _
" FROM PedidoInsumosFijosSemana" & _
" GROUP BY PedidoInsumosFijosSemana.CodCotizacion, PedidoInsumosFijosSemana.Semana" & _
" HAVING (((PedidoInsumosFijosSemana.CodCotizacion)=" & coti & ") AND ((PedidoInsumosFijosSemana.Semana)=" & sema & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PedidoInsumoFijoCotizacionSemana = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
PedidoInsumoFijoCotizacionSemana = 0

End Function

Function EstandarInsumoFijoCotizacion(coti As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT EstandarInsumosFijos.Registro, EstandarInsumosFijos.CodCotizacion, EstandarInsumosFijos.Cantidad" & _
" FROM EstandarInsumosFijos" & _
" WHERE (((EstandarInsumosFijos.CodCotizacion) = " & coti & "))" & _
" ORDER BY EstandarInsumosFijos.Registro DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
EstandarInsumoFijoCotizacion = rst("Cantidad")
rst.Close

Exit Function

Nulo:
EstandarInsumoFijoCotizacion = 0

End Function

Function EstandarInsumoFijoCotizacionxLocal(coti As Integer, loc As Integer) As Double
'
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT EstandarInsumosFijos.Registro, EstandarInsumosFijos.CodCotizacion, EstandarInsumosFijos.local, EstandarInsumosFijos.Cantidad" & _
" FROM EstandarInsumosFijos IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((EstandarInsumosFijos.CodCotizacion) = " & coti & ") AND ((EstandarInsumosFijos.local) = " & loc & "))" & _
" ORDER BY EstandarInsumosFijos.Registro DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
EstandarInsumoFijoCotizacionxLocal = rst("Cantidad")
rst.Close

Exit Function

Nulo:
EstandarInsumoFijoCotizacionxLocal = 0

End Function


Function PedidoInsumosLocalesDia(coti As Integer, fechi As Date) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT PedidoInsumosLocales.CodCotizacion, PedidoInsumosLocales.Fecha," & _
" Sum(PedidoInsumosLocales.Cantidad) AS SumaDeCantidad" & _
" FROM PedidoInsumosLocales" & _
" GROUP BY PedidoInsumosLocales.CodCotizacion, PedidoInsumosLocales.Fecha" & _
" HAVING (((PedidoInsumosLocales.CodCotizacion)=" & coti & ") AND ((PedidoInsumosLocales.Fecha)=#" & fechi & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PedidoInsumosLocalesDia = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
PedidoInsumosLocalesDia = 0

End Function

Function ContadorItemsPreingreso(prein As Long) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Sum(1) AS cont, SubPreIngresosPitaya.CodPreIngresoPitaya FROM SubPreIngresosPitaya" & _
" GROUP BY SubPreIngresosPitaya.CodPreIngresoPitaya HAVING (((SubPreIngresosPitaya.CodPreIngresoPitaya)=" & prein & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ContadorItemsPreingreso = rst("cont")
rst.Close

Exit Function

Nulo:
ContadorItemsPreingreso = 0

End Function


Function cantidadcambiosproductopreingreso(cotix As Integer, subprein As Long) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CambiosPreIngresosPitaya.CodCotizacion, CambiosPreIngresosPitaya.CodSubPreIngresoPitaya," & _
" Sum(CambiosPreIngresosPitaya.Cantidad) AS SumaDeCantidad" & _
" FROM CambiosPreIngresosPitaya" & _
" GROUP BY CambiosPreIngresosPitaya.CodCotizacion, CambiosPreIngresosPitaya.CodSubPreIngresoPitaya" & _
" HAVING (((CambiosPreIngresosPitaya.CodCotizacion)=" & cotix & ") AND ((CambiosPreIngresosPitaya.CodSubPreIngresoPitaya)=" & subprein & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadcambiosproductopreingreso = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadcambiosproductopreingreso = 0

End Function

Function cantidadcambiosproductopreingresomixed(cotix As Long, subprein As Long) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CambiosPreingresoSucursalMixed.CodCotizacion, CambiosPreingresoSucursalMixed.CodSubPreIngresoPitaya," & _
" Sum(CambiosPreingresoSucursalMixed.Cantidad) AS SumaDeCantidad" & _
" FROM CambiosPreingresoSucursalMixed" & _
" GROUP BY CambiosPreingresoSucursalMixed.CodCotizacion, CambiosPreingresoSucursalMixed.CodSubPreIngresoPitaya" & _
" HAVING (((CambiosPreingresoSucursalMixed.CodCotizacion)=" & cotix & ") AND ((CambiosPreingresoSucursalMixed.CodSubPreIngresoPitaya)=" & subprein & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadcambiosproductopreingresomixed = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cantidadcambiosproductopreingresomixed = 0

End Function

Function existencambiosdesucursal(prein As Long) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CambiosPreingresoSucursalMixed.CodPreIngresoPitaya, Sum(1) AS Expr1" & _
" FROM CambiosPreingresoSucursalMixed" & _
" GROUP BY CambiosPreingresoSucursalMixed.CodPreIngresoPitaya" & _
" HAVING (((CambiosPreingresoSucursalMixed.CodPreIngresoPitaya)=" & prein & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
existencambiosdesucursal = rst("Expr1")
rst.Close

Exit Function

Nulo:
existencambiosdesucursal = 0

End Function

Function statuscambiospreingresoporsucursal(prein As Long) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT AprobacionCambiosPreIngresoSucursal.CodPreIngresoPitaya, AprobacionCambiosPreIngresoSucursal.Status" & _
" FROM AprobacionCambiosPreIngresoSucursal" & _
" WHERE (((AprobacionCambiosPreIngresoSucursal.CodPreIngresoPitaya)=" & prein & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
statuscambiospreingresoporsucursal = rst("Status")
rst.Close

Exit Function

Nulo:
statuscambiospreingresoporsucursal = 0

End Function
