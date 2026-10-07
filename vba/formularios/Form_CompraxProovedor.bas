' ==========================================================
' Modulo  : Form_CompraxProovedor
' Tipo    : 100  |  Lineas: 178
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database

Private Sub Comando1030_Click()
DoCmd.OpenForm "DetalleCompraxProovedor"
Forms![DetalleCompraxProovedor]![cproov] = Me.CodProovedor
Forms![DetalleCompraxProovedor]![Texto548] = Me.Nombre
Forms![DetalleCompraxProovedor].Form.Requery
End Sub

Private Sub Comando324_Click()
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cod As Integer

'recorrer todos los pooovedores
miSQL = "SELECT Proovedores.Nombre, Proovedores.CodProovedor" & _
" FROM DBIngredientes INNER JOIN Proovedores ON DBIngredientes.ProovedorPrincipal = Proovedores.CodProovedor" & _
" GROUP BY Proovedores.Nombre, Proovedores.CodProovedor" & _
" HAVING ((Not (Proovedores.CodProovedor) = 221 And Not (Proovedores.CodProovedor) = 222 And Not (Proovedores.CodProovedor) = 223))" & _
" ORDER BY Proovedores.Nombre"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    cod = rst("CodProovedor")
    DoCmd.OpenForm "DetalleCompraxProovedor"
    Forms![DetalleCompraxProovedor]![cproov] = cod
    Forms![DetalleCompraxProovedor]![asemana] = Me.asemana
    Forms![DetalleCompraxProovedor]![crango] = Me.arango
    Call Forms("[DetalleCompraxProovedor]").Comando324_Click
    rst.MoveNext
Loop

MsgBox "Listas impresas"
Exit Sub

Nulo:
MsgBox "No se completo, volver a imprimir"

End Sub

Private Sub Comando92_Click()

If IsNull(Me.asemana) Then
    MsgBox "Ingresar numero de semana"
    Exit Sub
End If

If IsNull(Me.arango) Then
    MsgBox "Ingresar rango de compra"
    Exit Sub
End If

Dim proov As Integer
proov = Me.CodProovedor
'Crear nueva orden de compra
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO OrdenDeCompra(codproovedor, fechaorden, tipopago)" & _
" values (" & proov & ", #" & Date & "#, '" & DLookup("[TipoDePago]", "[Proovedores]", "[CodProovedor] = " & proov) & "')"
DoCmd.SetWarnings True

'Codigo de ultima orden de compra
Dim codordencompra As Long
codordencompra = ultimaordendecompra()

'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente, [DBIngredientes]![ProovedorPrincipal]=" & proov & " AS cond," & _
" CotiPrincipalDeIngrediente([DBIngredientes]![CodIngrediente]) AS Coti" & _
" FROM DBIngredientes" & _
" WHERE ((([DBIngredientes]![ProovedorPrincipal]= " & proov & ")<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst

' Inicializar variables para las fórmulas
Dim comprasSemanaCalculo As Double
Dim stockinicial As Double
Dim proyeccion As Double
Dim Conversion As Double
Dim canti As Double

Dim I As Integer
Dim Consumo As Double
Dim SI As Double
Dim pedi As Double

Dim totalSucursales As Long
totalSucursales = sucursalestotales()

Do While Not rst.EOF
        
    ' Calcular comprasSemanaCalculo para cada ítem
    comprasSemanaCalculo = DLookup("[ComprasSemana]", "[DBIngredientes]", "[CodIngrediente]='" & rst("CodIngrediente") & "'")

    ' Inicializar proyeccion
    proyeccion = 0
    
    ' Calcular proyeccion y stockInicial usando un bucle para todos los ítems
    For I = 1 To totalSucursales
        ' Calcular consumo, si y pedi para cada ítem
        Consumo = FRConsumoSemanalMaximoIngredienteLocal(rst("CodIngrediente"), Me.asemana, Me.arango, I)
        SI = StockIngredientelocal(rst("CodIngrediente"), Me.asemana, I) + StockCotizacionlocalconversionestandar(rst("CodIngrediente"), Me.asemana, I)
        pedi = absolutopositivo(Consumo * comprasSemanaCalculo - SI)
        
        ' Acumular en proyeccion y stockInicial
        proyeccion = proyeccion + pedi

    Next I
    
    ' Calcular conversion y cantidad
    Conversion = conversioncalculado(CotiPrincipalDeIngrediente(rst("CodIngrediente")), Me.asemana - 1)
    stockinicial = StockIngredientelocal(rst("CodIngrediente"), Me.asemana, 0) + StockCotizacionlocalconversionestandar(rst("CodIngrediente"), Me.asemana, 0)
    canti = absolutopositivo(proyeccion - stockinicial)
    canti = redondear_mas(canti / Conversion)
        
    If canti <> 0 Then
    
        ' Obtener el precio unitario del producto
        Dim precioUnitario As Double
        
        ' Consulta SQL para obtener el precio unitario
        Dim precioSQL As String
        precioSQL = "SELECT TOP 1 SubOrdenDeCompra.costounitario " & _
                    "FROM SubOrdenDeCompra " & _
                    "INNER JOIN OrdenDeCompra ON SubOrdenDeCompra.codordendecompra = OrdenDeCompra.codordendecompra " & _
                    "WHERE SubOrdenDeCompra.codcotizacion = " & rst("Coti") & " " & _
                    "ORDER BY OrdenDeCompra.fechaorden DESC;"
        
        ' Ejecutar la consulta y obtener el precio unitario
        Dim precioRS As DAO.Recordset
        Set precioRS = CurrentDb.OpenRecordset(precioSQL, dbOpenDynaset)
        
        If Not precioRS.EOF Then
            precioUnitario = precioRS("costounitario")
        Else
            ' Si no se encuentra un precio, establecerlo como cero
            precioUnitario = 0
        End If
        
        precioRS.Close
    
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO SubOrdenDeCompra (codordendecompra, codcotizacion, cantidadorden, destino, costounitario) " & _
        " values (" & codordencompra & ", " & rst("Coti") & ", " & canti & ", 0, " & precioUnitario & ")"
        DoCmd.SetWarnings True
        
    End If
    
    rst.MoveNext
Loop
rst.Close

' Insertar en SubOrdenDeCompra
'DoCmd.SetWarnings False
'DoCmd.RunSQL "INSERT INTO SubOrdenDeCompra (codordendecompra, codcotizacion, cantidadorden, destino, costounitario) " & _
'             "SELECT " & codordencompra & " AS codordendecompra, CotiPrincipalDeIngrediente([CodIngrediente]) AS codcotizacion, 0 AS cantidadorden, 0 AS destino, 0 AS costounitario " & _
'             "FROM DBIngredientes " & _
'             "WHERE ProovedorPrincipal = " & Me.CodProovedor
'DoCmd.SetWarnings True

DoCmd.OpenForm "OrdenDeCompraPlantilla"

[Forms]![OrdenDeCompraPlantilla]![ordenbusqueda] = codordencompra
[Forms]![OrdenDeCompraPlantilla]![afechaorden] = Date
[Forms]![OrdenDeCompraPlantilla]![aproovedor] = Me.CodProovedor
Forms![OrdenDeCompraPlantilla]![tipopago].Value = DLookup("TipoDePago", "Proovedores", "CodProovedor = " & Me.CodProovedor)
[Forms]![OrdenDeCompraPlantilla].Requery
Exit Sub

Nulo:
MsgBox "No se ha creado Orden de Compra"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
