' ==========================================================
' Modulo  : Rendimientos
' Tipo    : 1
' Lineas  : 158
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database
Function ConversionEstandar(coti As Long) As Double
Dim convedefecto As Double
convedefecto = DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & coti)
If convedefecto = 0 Then
    ConversionEstandar = DLookup("[ConversionEstandar]", "[Cotizaciones]", "[CodCotizacion]=" & coti)
Else
    ConversionEstandar = convedefecto
End If

Exit Function

Nulo:
ConversionEstandar = 0
End Function

Function conversioncalculado(coti As Integer, semana As Integer) As Double
Dim conv As Double
'SCOnversion o rendimientos de producto acorde  a historial procesado de la semana mencionada
conv = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti)

If conv = 0 Then 'Conversion Calculada

    On Error GoTo Nulo
    Dim rst As DAO.Recordset
    Dim miSQL As String

    'valor de procesamiento total y cantidad procesada
    miSQL = "SELECT numerosemana([SubPorcionamiento]![Fecha])<=" & semana & " AS [cond 3]," & _
    " numerosemana([SubPorcionamiento]![Fecha]) AS semana, SubPorcionamiento.Procedencia," & _
    " Sum(porcionadoselladoxsubpor([SubPorcionamiento]![CodSubPorcionamiento])) AS Final," & _
    " Sum(SubPorcionamiento.Cantidad) AS SumaDeCantidad" & _
    " FROM SubPorcionamiento" & _
    " GROUP BY numerosemana([SubPorcionamiento]![Fecha])<=" & semana & ", numerosemana([SubPorcionamiento]![Fecha])," & _
    " SubPorcionamiento.Procedencia" & _
    " HAVING (((numerosemana([SubPorcionamiento]![Fecha]) <= " & semana & ") = True) And ((SubPorcionamiento.Procedencia) = " & coti & "))" & _
    " ORDER BY numerosemana([SubPorcionamiento]![Fecha]) DESC"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    rst.MoveFirst
    conversioncalculado = rst("Final") / rst("SumaDeCantidad")  'COnversion de Insumo
    rst.Close

Else 'COnversion natural

    conversioncalculado = conv

End If

Exit Function

Nulo:
conversioncalculado = conv


End Function

Function conversioncalculadoxlocal(coti As Integer, semana As Integer, locx As Integer) As Double
Dim conv As Double
'SCOnversion o rendimientos de producto acorde  a historial procesado de la semana mencionada
conv = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti)

If conv = 0 Then 'Conversion Calculada

    On Error GoTo Nulo
    Dim rst As DAO.Recordset
    Dim miSQL As String

    'valor de procesamiento total y cantidad procesada
    miSQL = "SELECT numerosemana([SubPorcionamiento" & locx & "]![Fecha])<=" & semana & " AS [cond 3]," & _
    " numerosemana([SubPorcionamiento" & locx & "]![Fecha]) AS semana, SubPorcionamiento" & locx & ".Procedencia," & _
    " Sum(porcionadoselladoxsubporxlocal([SubPorcionamiento" & locx & "]![CodSubPorcionamiento]," & locx & ")) AS Final," & _
    " Sum(SubPorcionamiento" & locx & ".Cantidad) AS SumaDeCantidad" & _
    " FROM SubPorcionamiento" & locx & " IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
    " GROUP BY numerosemana([SubPorcionamiento" & locx & "]![Fecha])<=" & semana & ", numerosemana([SubPorcionamiento" & locx & "]![Fecha])," & _
    " SubPorcionamiento" & locx & ".Procedencia" & _
    " HAVING (((numerosemana([SubPorcionamiento" & locx & "]![Fecha]) <= " & semana & ") = True)" & _
    " And ((SubPorcionamiento" & locx & ".Procedencia) = " & coti & "))" & _
    " ORDER BY numerosemana([SubPorcionamiento" & locx & "]![Fecha]) DESC"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    rst.MoveFirst
    conversioncalculadoxlocal = rst("Final") / rst("SumaDeCantidad")  'COnversion de Insumo
    rst.Close

Else 'COnversion natural

    conversioncalculadoxlocal = conv

End If

Exit Function

Nulo:
conversioncalculadoxlocal = conv


End Function

Function ConversionCalculadoGlobal(coti As Integer, semana As Integer) As Double
Dim conv As Double
'SCOnversion o rendimientos de producto acorde  a historial procesado de la semana mencionada
conv = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti)

If conv = 0 Then 'Conversion Calculada

    On Error GoTo Nulo
    Dim rst As DAO.Recordset
    Dim miSQL As String

    'valor de procesamiento total y cantidad procesada
    miSQL = "SELECT numerosemana([SubPorcionamiento]![Fecha])<=" & semana & " AS [cond 3]," & _
    " numerosemana([SubPorcionamiento]![Fecha]) AS semana, SubPorcionamiento.Procedencia," & _
    " Sum(porcionadoselladoxsubporxlocal([SubPorcionamiento]![CodSubPorcionamiento],[SubPorcionamiento]![local])) AS Final," & _
    " Sum(SubPorcionamiento.Cantidad) AS SumaDeCantidad" & _
    " FROM SubPorcionamiento IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
    " GROUP BY numerosemana([SubPorcionamiento]![Fecha])<=" & semana & ", numerosemana([SubPorcionamiento]![Fecha])," & _
    " SubPorcionamiento.Procedencia" & _
    " HAVING (((numerosemana([SubPorcionamiento]![Fecha]) <= " & semana & ") = True)" & _
    " And ((SubPorcionamiento.Procedencia) = " & coti & "))" & _
    " ORDER BY numerosemana([SubPorcionamiento]![Fecha]) DESC"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    ConversionCalculadoGlobal = rst("Final") / rst("SumaDeCantidad")  'COnversion de Insumo
    rst.Close

Else 'COnversion natural

    ConversionCalculadoGlobal = conv

End If

Exit Function

Nulo:
ConversionCalculadoGlobal = conv


End Function

Function OperarioAzar() As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Operarios.CodOperario, Operarios.Operativo, Operarios.Ciudad" & _
" FROM Operarios" & _
" WHERE (((Operarios.Operativo) = True) And ((Operarios.Ciudad) = ciudadsistema()))" & _
" ORDER BY Operarios.CodOperario"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
OperarioAzar = rst("CodOperario")
rst.Close

Exit Function

Nulo:
OperarioAzar = 1

End Function

