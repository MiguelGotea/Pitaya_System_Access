' ==========================================================
' Modulo  : Productos
' Tipo    : 1  |  Lineas: 789
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database
Function CotiPrincipalDeIngrediente(ingr As String) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.Conversion, Cotizaciones.CodIngrediente, Cotizaciones.Prioridad," & _
" Cotizaciones.Subproducto, IIf([Cotizaciones]![Marca]='Almacen Global',1,0) AS condi, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones" & _
" WHERE (((Cotizaciones.CodIngrediente)='" & ingr & "') AND ((Cotizaciones.Prioridad)<>0)" & _
" AND ((Cotizaciones.Subproducto)=0) AND ((IIf([Cotizaciones]![Marca]='Almacen Global',1,0))=0))" & _
" ORDER BY Cotizaciones.Conversion"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
CotiPrincipalDeIngrediente = rst("CodCotizacion")
rst.Close
Exit Function
Nulo:
CotiPrincipalDeIngrediente = 0
End Function

Function CotisNoPorcionPrioridadVigente(cotid As Long) As Boolean
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.Prioridad, Cotizaciones.Subproducto," & _
" EsAlmacenGlobalDePorcion([Cotizaciones]![CodCotizacion]) AS AG, Cotizaciones.Descontinuado," & _
" 1 AS existe, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones" & _
" WHERE (((Cotizaciones.Prioridad)<>0) AND ((Cotizaciones.Subproducto)=0)" & _
" AND ((EsAlmacenGlobalDePorcion([Cotizaciones]![CodCotizacion]))=0)" & _
" AND ((Cotizaciones.Descontinuado)=0) AND ((Cotizaciones.CodCotizacion)=" & cotid & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
If rst("existe") = 1 Then
    CotisNoPorcionPrioridadVigente = True
Else
    CotisNoPorcionPrioridadVigente = False
End If

rst.Close

Exit Function
Nulo:
CotisNoPorcionPrioridadVigente = False
End Function

Function EsAlmacenGlobalDePorcion(cotid As Long) As Boolean
On Error GoTo Nulo
Dim Marca As String
Marca = DLookup("[Marca]", "[Cotizaciones]", "[CodCotizacion]=" & cotid)

If Marca = "Almacen Global" Then
    EsAlmacenGlobalDePorcion = True
Else
    EsAlmacenGlobalDePorcion = False
End If

Exit Function
Nulo:
EsAlmacenGlobalDePorcion = False
End Function

Function porciondeingredientevigentenomostrador(ingr As String) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [DBBatidos]![CodGrupo]<>7 AS Expr1, DBBatidos.Vigencia, SubReceta.codporcion, SubReceta.CodIngrediente" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY [DBBatidos]![CodGrupo]<>7, DBBatidos.Vigencia, SubReceta.codporcion, SubReceta.CodIngrediente" & _
" HAVING ((([DBBatidos]![CodGrupo]<>7)<>0) AND ((DBBatidos.Vigencia)<>0) AND ((SubReceta.codporcion) Is Not Null)" & _
" AND ((SubReceta.CodIngrediente)='" & ingr & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
porciondeingredientevigentenomostrador = rst("codporcion")
rst.Close
Exit Function
Nulo:
porciondeingredientevigentenomostrador = 0
End Function

Function CotiDirectoDeIngrediente(ingr As String) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.CodIngrediente, Cotizaciones.Conversion, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones WHERE (((Cotizaciones.CodIngrediente)='" & ingr & "') AND ((Cotizaciones.Conversion)=1))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CotiDirectoDeIngrediente = rst("CodCotizacion")
rst.Close
Exit Function
Nulo:
CotiDirectoDeIngrediente = 0
End Function



Function CotiPrincipalProdCompraVenta(bat As String) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.CodBatido, SubReceta.InsumoClave," & _
" IIf(IsNull([SubReceta]![codporcion]),cotidirectodeingrediente([Subreceta]![CodIngrediente]),[SubReceta]![codporcion]) AS cond" & _
" FROM (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido" & _
" WHERE (((DBBatidos.CodBatido)='" & bat & "') AND ((SubReceta.InsumoClave)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CotiPrincipalProdCompraVenta = rst("cond")
rst.Close
Exit Function
Nulo:
CotiPrincipalProdCompraVenta = 0
End Function
Function IngrePrincipalProdCompraVenta(bat As String) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.CodBatido, DBIngredientes.Tipo, SubReceta.CodIngrediente" & _
" FROM (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido" & _
" WHERE (((DBBatidos.CodBatido)='" & bat & "') AND ((DBIngredientes.Tipo)<>'Empaque'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngrePrincipalProdCompraVenta = rst("CodIngrediente")
rst.Close
Exit Function
Nulo:
IngrePrincipalProdCompraVenta = 0
End Function

Function ListaCodigosBarra(cotit As Long) As String
'lista en una linea decodos de barrade una coti
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer

'todos los codigos de barra de ese codigo cotiacion
miSQL = "SELECT CodigoBarraCotizacion.CodCotizacion, CodigoBarraCotizacion.CodigoBarra" & _
" FROM CodigoBarraCotizacion WHERE (((CodigoBarraCotizacion.CodCotizacion)=" & cotit & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst
ListaCodigosBarra = ""

For I = 1 To canti
    If I = 1 Then
        ListaCodigosBarra = rst("CodigoBarra")
    Else
        ListaCodigosBarra = ListaCodigosBarra & ", " & rst("CodigoBarra")
    End If
    rst.MoveNext
Next I

rst.Close

Exit Function

Nulo:
ListaCodigosBarra = ""

End Function

Function ListaCodigosBarraProductos(cotit As String) As String
'lista en una linea decodos de barrade una  producto de venta
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer

'todos los codigos de barra de ese codigo producto venta
miSQL = "SELECT CodigoBarraBatidos.CodBatido, CodigoBarraBatidos.CodigoBarra" & _
" FROM CodigoBarraBatidos WHERE (((CodigoBarraBatidos.CodBatido)='" & cotit & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst
ListaCodigosBarraProductos = ""

For I = 1 To canti
    If I = 1 Then
        ListaCodigosBarraProductos = rst("CodigoBarra")
    Else
        ListaCodigosBarraProductos = ListaCodigosBarraProductos & ", " & rst("CodigoBarra")
    End If
    rst.MoveNext
Next I

rst.Close

Exit Function

Nulo:
ListaCodigosBarraProductos = ""

End Function

Function TipoOrigenProductoCompraVenta(bat As String) As Integer
'0 lleva porcion
'-1 es igrediente
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.CodBatido, DBBatidos.Vigencia, DBBatidos.CompraVenta, DBIngredientes.Tipo," & _
" IsNull([SubReceta]![codporcion]) AS llevaporcion, SubReceta.CodIngrediente, SubReceta.codporcion" & _
" FROM (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido" & _
" WHERE (((DBBatidos.CodBatido)='" & bat & "') AND ((DBBatidos.Vigencia)<>0) AND ((DBBatidos.CompraVenta)<>0)" & _
" AND ((DBIngredientes.Tipo)<>'Empaque'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TipoOrigenProductoCompraVenta = rst("llevaporcion")
rst.Close
Exit Function
Nulo:
TipoOrigenProductoCompraVenta = -1
'por defecto la cantidad de ingrediente
End Function


Function IngredienteContieneReceta(bat As String, ingre As String) As Integer
'1 lleva ingrediente
'0 no tiene igrediente
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.CodBatido, SubReceta.CodIngrediente, 1 AS indi" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" WHERE (((DBBatidos.CodBatido)='" & bat & "') AND ((SubReceta.CodIngrediente)='" & ingre & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngredienteContieneReceta = rst("indi")
rst.Close
Exit Function
Nulo:
IngredienteContieneReceta = 0
End Function

Function factorabastecimiento(codal As Integer, seco As Double, refr As Double, conge As Double) As Double

Select Case codal
    Case 1 'congelado
        factorabastecimiento = conge
    Case 2 'refrigerado
        factorabastecimiento = refr
    Case 3 'seco
        factorabastecimiento = seco
    Case Else
        factorabastecimiento = 0
End Select
End Function



Function ListaInsumoProductoVenta(cotit As String) As String
'lista en una linea de todos los inusmo de producto venta
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer

miSQL = "SELECT DBIngredientes.Nombre, SubReceta.CodBatido" & _
" FROM DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente" & _
" WHERE (((SubReceta.CodBatido)='" & cotit & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst
ListaInsumoProductoVenta = ""

For I = 1 To canti
    ListaInsumoProductoVenta = ListaInsumoProductoVenta & " , " & rst("Nombre")
    rst.MoveNext
Next I

rst.Close

Exit Function

Nulo:
ListaInsumoProductoVenta = ""

End Function

Function grupoproductovigente(cotit As Long) As Boolean
'lista en una linea de todos los inusmo de producto venta
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.Tipo, DBIngredientes.Nombre, [DBBatidos]![CodGrupo]=7 AS Expr2, DBBatidos.Vigencia, SubReceta.codporcion" & _
" FROM (SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" GROUP BY DBIngredientes.Tipo, DBIngredientes.Nombre, [DBBatidos]![CodGrupo]=7, DBBatidos.Vigencia, SubReceta.codporcion" & _
" HAVING ((([DBBatidos]![CodGrupo] = 7) = False) And ((DBBatidos.Vigencia) = True) And ((SubReceta.codporcion) = " & cotit & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If IsNull(rst("Nombre")) Or rst("Nombre") = "" Then
    grupoproductovigente = False
Else
    grupoproductovigente = True
End If
rst.Close

Exit Function

Nulo:
grupoproductovigente = False

End Function


Function CotiNoPorcionFiltro(cotit As Long) As Boolean
'verifica si una cotizacion esta dentro de porciones vigentes
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [DBBatidos]![CodGrupo]=7 AS Expr1, SubReceta.codporcion," & _
" CotiPrincipalDeIngrediente([SubReceta]![CodIngrediente]) AS coti, DBIngredientes.Nombre" & _
" FROM (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE ((([DBBatidos]![CodGrupo]=7)=0) AND ((SubReceta.codporcion) Is Null)" & _
" AND ((CotiPrincipalDeIngrediente([SubReceta]![CodIngrediente]))=" & cotit & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If IsNull(rst("Nombre")) Or rst("Nombre") = "" Then
    CotiNoPorcionFiltro = False
Else
    CotiNoPorcionFiltro = True
End If
rst.Close

Exit Function

Nulo:
CotiNoPorcionFiltro = False

End Function

Function InsumoMezclaPorcion(cotiz As Integer) As Integer
'Si el ingrediente es parte de una mezcla de porciones
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.codporcion, Cotizaciones.MezclaPorcion" & _
" FROM SubReceta INNER JOIN Cotizaciones ON SubReceta.codporcion = Cotizaciones.CodCotizacion" & _
" WHERE (((SubReceta.codporcion)=" & cotiz & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
If rst("MezclaPorcion") Then
    InsumoMezclaPorcion = 1
Else
    InsumoMezclaPorcion = 0
End If

rst.Close
Exit Function

Nulo:
InsumoMezclaPorcion = 0
End Function
Function ExisteMezcla(CodCotizacionMezcla As Integer) As Integer
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String
Dim existeRelacion As Integer

' Inicializar el valor predeterminado
existeRelacion = 0

' Construir la consulta SQL
strSQL = "SELECT TOP 1 CodMezclaPorciones FROM MezclaPorciones WHERE CodCotizacionMezcla = " & CodCotizacionMezcla
Set db = CurrentDb
Set rs = db.OpenRecordset(strSQL)

' Comprobar si se encontró algún registro
If Not rs.EOF Then
    ' Si hay registros, significa que existe relación
    existeRelacion = 1
End If

rs.Close
Set rs = Nothing
Set db = Nothing
ExisteMezcla = existeRelacion

Exit Function
ErrorHandler:
ExisteMezcla =  0
End Function

Function PorcionDentroDeMezcla(CodCotizacionPorcion As Long) As Integer

On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String

' Construir la consulta SQL
strSQL = "SELECT MezclaPorciones.CodCotizacionPorcion, 1 AS contador" & _
" FROM MezclaPorciones WHERE (((MezclaPorciones.CodCotizacionPorcion)=" & CodCotizacionPorcion & "))"

Set db = CurrentDb
Set rs = db.OpenRecordset(strSQL)

If rs("contador") = 1 Then
    PorcionDentroDeMezcla = 1
Else
    PorcionDentroDeMezcla = 0
    
End If

rs.Close
Set rs = Nothing
Set db = Nothing

Exit Function
ErrorHandler:
PorcionDentroDeMezcla =  0

End Function


Function cantidadProductosMezcla(CodCotizacionMezcla As Integer) As Integer

On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String
Dim cantidadSubproductos As Integer

' Inicializar el valor predeterminado
cantidadSubproductos = 0

' Construir la consulta SQL
strSQL = "SELECT COUNT(*) AS CantidadSubproductos FROM MezclaPorciones WHERE CodCotizacionMezcla = " & CodCotizacionMezcla
Set db = CurrentDb
Set rs = db.OpenRecordset(strSQL)

' Obtener la cantidad de subproductos
If Not rs.EOF Then
    cantidadSubproductos = rs!cantidadSubproductos
End If

rs.Close
Set rs = Nothing
Set db = Nothing
cantidadProductosMezcla = cantidadSubproductos

Exit Function
ErrorHandler:
cantidadProductosMezcla =  0

End Function


Function ProductoConsumePresentacionPorcion(ingref As String) As Integer
'0: no se usa en ningunproducto como porcione
'1: usa presentacion de porcion en alguna receta
On Error GoTo Nulo

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String

' Construir la consulta SQL
strSQL = "SELECT DBBatidos.Vigencia, DBIngredientes.CodIngrediente, Sum(IIf(IsNull([SubReceta]![codporcion]),0,1)) AS Expr1" & _
" FROM (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" GROUP BY DBBatidos.Vigencia, DBIngredientes.CodIngrediente" & _
" HAVING (((DBBatidos.Vigencia)<>0) AND ((DBIngredientes.CodIngrediente)='" & ingref & "'))"

Set db = CurrentDb
Set rs = db.OpenRecordset(strSQL)

If rs("Expr1") > 0 Then
    ProductoConsumePresentacionPorcion = 1
Else
    ProductoConsumePresentacionPorcion = 0
End If

rs.Close
Set rs = Nothing
Set db = Nothing

Exit Function
Nulo:
ProductoConsumePresentacionPorcion =  0

End Function

Function ProductoConsumePresentacionNoPorcion(ingref As String) As Integer
'0: no se usa en ningunproducto como porcione
'1: usa presentacion de porcion en alguna receta
On Error GoTo Nulo

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String

' Construir la consulta SQL
strSQL = "SELECT DBBatidos.Vigencia, DBIngredientes.CodIngrediente, Sum(IIf(IsNull([SubReceta]![codporcion]),1,0)) AS Expr1" & _
" FROM (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" GROUP BY DBBatidos.Vigencia, DBIngredientes.CodIngrediente" & _
" HAVING (((DBBatidos.Vigencia)<>0) AND ((DBIngredientes.CodIngrediente)='" & ingref & "'))"

Set db = CurrentDb
Set rs = db.OpenRecordset(strSQL)

If rs("Expr1") > 0 Then
    ProductoConsumePresentacionNoPorcion = 1
Else
    ProductoConsumePresentacionNoPorcion = 0
End If

rs.Close
Set rs = Nothing
Set db = Nothing

Exit Function
Nulo:
ProductoConsumePresentacionNoPorcion =  0

End Function

Function yaexistecodigobatido(codib As String) As Integer
'1 si existe 0 si no existe codigo de batido
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.CodBatido, 1 AS existe FROM DBBatidos WHERE (((DBBatidos.CodBatido)='" & codib & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
yaexistecodigobatido = rst("existe")

rst.Close
Exit Function

Nulo:
yaexistecodigobatido = 0
End Function

Public Sub actualizarfiltroporcionesvigentes()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM FiltroPorcionesVigentes"
DoCmd.SetWarnings True

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim cot As Long
Dim afech As Date

miSQL = "SELECT [DBBatidos]![CodGrupo]<>7 AS Expr2, SubReceta.codporcion, PorcionDentroDeMezcla([SubReceta]![codporcion]) AS mezcla," & _
" Grupos.control, DLookUp('[Descontinuado]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]) AS Desontinuado, DBIngredientes.Vigente" & _
" FROM ((SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente)" & _
" INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY [DBBatidos]![CodGrupo]<>7, SubReceta.codporcion," & _
" PorcionDentroDeMezcla([SubReceta]![codporcion]), Grupos.control," & _
" DLookUp('[Descontinuado]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]), DBIngredientes.Vigente" & _
" HAVING ((([DBBatidos]![CodGrupo]<>7)<>0) AND ((SubReceta.codporcion) Is Not Null)" & _
" AND ((PorcionDentroDeMezcla([SubReceta]![codporcion]))=0) AND ((Grupos.control)<>0)" & _
" AND ((DLookUp('[Descontinuado]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion]))=0)" & _
" AND ((DBIngredientes.Vigente)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = Now

For I = 1 To cantli
    cot = rst("codporcion")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [FiltroPorcionesVigentes](CodCotizacion, FechaActualizacion)" & _
    " values (" & cot & ", #" & afech & "#)"
    DoCmd.SetWarnings True
    rst.MoveNext
Next I
rst.Close

Exit Sub

AgainAgain:
rst.Close
End Sub
Public Sub actualizarfiltronoporcionesvigentes()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM FiltroNoPorcionesVigentes"
DoCmd.SetWarnings True

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim cot As String
Dim afech As Date

miSQL = "SELECT DBBatidos.Vigencia, [DBBatidos]![CodGrupo]=7 AS nomostradorcombos," & _
" SubReceta.codporcion, Grupos.control, SubReceta.CodIngrediente" & _
" FROM (((DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente)" & _
" INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY DBBatidos.Vigencia, [DBBatidos]![CodGrupo]=7, SubReceta.codporcion, Grupos.control, SubReceta.CodIngrediente" & _
" HAVING (((DBBatidos.Vigencia)<>0) AND (([DBBatidos]![CodGrupo]=7)=0)" & _
" AND ((SubReceta.codporcion) Is Null) AND ((Grupos.control)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = Now

For I = 1 To cantli
    cot = rst("CodIngrediente")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [FiltroNoPorcionesVigentes](CodIngrediente, FechaActualizacion)" & _
    " values ('" & cot & "', #" & afech & "#)"
    DoCmd.SetWarnings True
    rst.MoveNext
Next I
rst.Close

Exit Sub

AgainAgain:
rst.Close
End Sub

Public Sub actualizarfiltromostradorvigentes()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM FiltroMostradorVigentes"
DoCmd.SetWarnings True

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim cot As Long
Dim ingr As String
Dim atip As Integer
Dim afech As Date

miSQL = "SELECT DBBatidos.Vigencia, DBBatidos.CodGrupo, SubReceta.InsumoClave, SubReceta.codporcion," & _
" SubReceta.CodIngrediente, IIf(IsNull([SubReceta]![codporcion]),1,2) AS tipo" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY DBBatidos.Vigencia, DBBatidos.CodGrupo, SubReceta.InsumoClave, SubReceta.codporcion," & _
" SubReceta.CodIngrediente, IIf(IsNull([SubReceta]![codporcion]),1,2)" & _
" HAVING (((DBBatidos.Vigencia)<>0) AND ((DBBatidos.CodGrupo)=7) AND ((SubReceta.InsumoClave)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = Now

For I = 1 To cantli
    If rst("tipo") = 1 Then
        ingr = rst("CodIngrediente")
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO [FiltroMostradorVigentes](Tipo, CodIngrediente, CodCotizacion, FechaActualizacion)" & _
        " values (1, '" & ingr & "', 0, #" & afech & "#)"
        DoCmd.SetWarnings True
        rst.MoveNext
    Else
        cot = rst("codporcion")
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO [FiltroMostradorVigentes](Tipo, CodIngrediente, CodCotizacion, FechaActualizacion)" & _
        " values (2, ' ', " & cot & ", #" & afech & "#)"
        DoCmd.SetWarnings True
        rst.MoveNext
    End If
    
Next I
rst.Close

Exit Sub

AgainAgain:
rst.Close
End Sub
Public Sub actualizarfiltromateriaprimavigentes()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM FiltroMateriaPrimaVigentes"
DoCmd.SetWarnings True

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim ingrex As String
Dim afech As Date

miSQL = "SELECT TiposVariables.Orden, DBIngredientes.Nombre, Grupos.control, TiposVariables.Control," & _
" DBBatidos.Vigencia, IIf(IsNull([SubReceta]![codporcion]),0,ExisteMezcla([SubReceta]![codporcion])) AS mezcla, SubReceta.CodIngrediente" & _
" FROM (((DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente)" & _
" INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo) INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" GROUP BY TiposVariables.Orden, DBIngredientes.Nombre, Grupos.control, TiposVariables.Control, DBBatidos.Vigencia," & _
" IIf(IsNull([SubReceta]![codporcion]),0,ExisteMezcla([SubReceta]![codporcion])), SubReceta.CodIngrediente" & _
" HAVING (((Grupos.Control) = True) And ((TiposVariables.Control) = True) And ((DBBatidos.Vigencia) = True)" & _
" And ((IIf(IsNull([SubReceta]![codporcion]), 0, ExisteMezcla([SubReceta]![codporcion]))) = 0))" & _
" ORDER BY TiposVariables.Orden, DBIngredientes.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = Now

For I = 1 To cantli
    ingrex = rst("CodIngrediente")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [FiltroMateriaPrimaVigentes](CodIngrediente, FechaActualizacion)" & _
    " values ('" & ingrex & "', #" & afech & "#)"
    DoCmd.SetWarnings True
    rst.MoveNext
Next I
rst.Close

Exit Sub

AgainAgain:
rst.Close
End Sub
Function cantidadingredientesrecetavisibles(codib As String) As Integer
'1 si existe 0 si no existe codigo de batido
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))) AS visibles," & _
" SubReceta.CodBatido, Sum(1) AS Expr1" & _
" FROM TipoIngredientesReceta INNER JOIN SubReceta ON TipoIngredientesReceta.CodTipoIngredientesReceta = SubReceta.Tipo" & _
" GROUP BY IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))), SubReceta.CodBatido" & _
" HAVING (((IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))))=0)" & _
" AND ((SubReceta.CodBatido)='" & codib & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadingredientesrecetavisibles = rst("Expr1")

rst.Close
Exit Function

Nulo:
cantidadingredientesrecetavisibles = 0
End Function

Function cantidadingredientesmaximorecetavisiblespedido(pedix As Long, Tipo As Integer) As Integer

'tipo 1 batidos , tipo 2 bowl y waffle
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim buscax As String

If Tipo = 1 Then
    buscax = "(Grupos.tipo) = 'Batido' Or (Grupos.tipo) = 'Limonada'"
Else
    buscax = "(Grupos.tipo) = 'Bowl' Or (Grupos.tipo) = 'Waffles' Or (Grupos.tipo) = 'Parfait'"
End If

miSQL = "SELECT cantidadingredientesrecetavisibles([SubPedido]![CodBatido]) AS maxim, SubPedido.CodPedido, Grupos.Tipo" & _
" FROM Grupos INNER JOIN (SubPedido INNER JOIN DBBatidos" & _
" ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" ON Grupos.CodGrupo = DBBatidos.CodGrupo" & _
" WHERE (((SubPedido.CodPedido) = " & pedix & ")" & _
" And (" & buscax & "))" & _
" ORDER BY cantidadingredientesrecetavisibles([SubPedido]![CodBatido]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadingredientesmaximorecetavisiblespedido = rst("maxim")

rst.Close
Exit Function

Nulo:
cantidadingredientesmaximorecetavisiblespedido = 0
End Function

