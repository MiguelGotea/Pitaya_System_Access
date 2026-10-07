' ==========================================================
' Modulo  : FormulasPedidos
' Tipo    : 1  |  Lineas: 2392
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database
Sub actualizarmontoguardadopedidosdia(fechi As Date)
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Recorrer todas las notas del pedido del dia
miSQL = "SELECT NotaDePedido.CodPedido, NotaDePedido.Fecha FROM NotaDePedido WHERE (((NotaDePedido.Fecha)=#" & fechi & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst

Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE NotaDePedido SET NotaDePedido.TotalGuardado = " & MontoPedido(rst("CodPedido")) & _
    " WHERE NotaDePedido.CodPedido = " & rst("CodPedido")
    DoCmd.SetWarnings True
    
    Call ActualizarPuntosSubPedidoDePedido(rst("CodPedido"))
    rst.MoveNext
Loop

Exit Sub

Nulo:
End Sub
Function LoopCocina(codigo As Long) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cod As Long

'Hallar la lista de subpedidos
miSQL = "SELECT SubPedido.CodSubPedido, SubPedido.CodPedido, Grupos.Imprimible" & _
" FROM (SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" WHERE (((SubPedido.CodPedido)=" & codigo & ") AND ((Grupos.Imprimible)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst

Do While Not rst.EOF
    cod = rst("CodSubPedido")
    
    'VISTA EN PANTALLA
    'DoCmd.OpenReport "Cocina", acViewReport, , "CodSubPedido = " & cod & ""
    'MsgBox "SIGUIENTE"
    'DoCmd.Close acReport, "Cocina"
    
    'IMPRESO
    DoCmd.OpenReport "Cocina", acViewNormal, , "CodSubPedido = " & cod & ""
    
    rst.MoveNext
Loop

Exit Function

Nulo:
LoopCocina = 0

End Function

Function usovidrio(codigo As Long) As String
' verificar si es vidrio o no la nota de pedido solicita codigo de nota de pedido

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim opt As Integer

'Hallar si es vidrio o no
miSQL = "SELECT NotaDePedido.CodPedido, CInt([Modalidad]) AS Modo" & _
" FROM NotaDePedido WHERE (NotaDePedido.CodPedido=" & codigo & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If rst("Modo") = 0 Then
    usovidrio = "Plastico"
Else
    usovidrio = "Vidrio"
End If


Exit Function

Nulo:
usovidrio = ""

End Function


Function PrecioReal(cod As Long) As Double
' PRECIO REAL UNITARIO , cod = codigo subpedido

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim codpro As Integer
Dim mont As Double

'Hallar el precio real de un producto con descuento Y CONSIDERANDO SI NO ESTA ANULADO
miSQL = "SELECT SubPedido.CodSubPedido, SubPedido.CodPromocion, SubPedido.CodBatido, [DBPromociones]![RatioDescuento]*[DBBatidos]![Precio] AS MontoReal," & _
" NotaDePedido.Anulado, DBBatidos.Medida, DBBatidos.CodGrupo" & _
" FROM NotaDePedido" & _
" INNER JOIN (DBBatidos INNER JOIN (DBPromociones INNER JOIN SubPedido ON DBPromociones.CodPromocion = SubPedido.CodPromocion)" & _
" ON DBBatidos.CodBatido = SubPedido.CodBatido) ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" WHERE (SubPedido.CodSubPedido=" & cod & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
codpro = rst("CodPromocion")
mont = rst("MontoReal")

If rst("Anulado") = 0 Then

    Select Case codpro
        Case 4                    ' Antiguo agranda batid
        PrecioReal = mont - 9
        Case 12                    ' Estreno Energy Green
        PrecioReal = mont - 10
        Case 24                    ' Agranda batido HASTA OCTUBRE 2017
        PrecioReal = mont - 8
        Case 37                    ' Agranda batido
        PrecioReal = mont - 9
        Case 76                    ' Agranda batido
        PrecioReal = mont - 10
        Case 56                    ' Agranda batido jueves
        PrecioReal = mont - 9
        Case 77                    ' Agranda batido jueves
        PrecioReal = mont - 10
        Case 30                    ' Kid a gigantona
        PrecioReal = mont - 19
        Case 42                    ' 15 PESOS REBAJA
        PrecioReal = mont - 15
        Case 20                    ' 2 por 100
        PrecioReal = 50
        Case 60                    ' Segundo Mediano a C$39
        PrecioReal = 39
        Case 61                    ' Segundo Grande a C$47
        PrecioReal = 47
        Case 72                    ' 8 cordobas producto compra venta
        PrecioReal = mont - 8
        Case 74                    ' 10 cordobas producto compra venta
        PrecioReal = mont - 10
        Case 64                    ' 10 pesos menos presentando volante
        PrecioReal = mont - 10
        Case 84                    ' 10 pesos menos si viene x influencer
        PrecioReal = mont - 10
        
        Case 116                   '2do waffle especial a 99 cordobas
        PrecioReal = 99
        Case 118                   '2 batidos especiales de 16oz x 139 cordobas
        PrecioReal = 69.5
        Case 119                   '2 Batidos con proteina de 16oz x 159 cordobas
        PrecioReal = 79.5
        Case 121                   '2 batidos medianos 16oz por 179
        PrecioReal = 89.5
        Case 141                   '2 batidos medianos 16oz por 200
        PrecioReal = 100
        Case 194                   '2 batidos medianos 16oz por 200
        PrecioReal = 107.5
        Case 217                   '2 batidos medianos 16oz por 240
        PrecioReal = 120
        Case 190                   '2 batidos especiale 16oz por 179
        PrecioReal = 89.5
        Case 215                   '2 batidos especiale 16oz por 198
        PrecioReal = 99
        Case 236                   '2 batidos especiale 16oz por 198
        PrecioReal = 90
        Case 237                   '2 batidos especiale 16oz por 198
        PrecioReal = 105
        Case 142                  '2 batidos con proteina  16oz por 179
        PrecioReal = 89.5
        Case 193                  '2 batidos con proteina  16oz por 179
        PrecioReal = 94.5
        Case 216                  '2 batidos con proteina  16oz por 210
        PrecioReal = 105
        Case 238                  '2 batidos con proteina  16oz por 210
        PrecioReal = 110
        Case 144                  '2 bowl dragon por 250
        PrecioReal = 125
        Case 162                  '2 bowl dragon por 250
        PrecioReal = 149.5
        Case 145                  '2 mediano clasico x 149
        PrecioReal = 74.5
        Case 192                  '2 mediano clasico x 165
        PrecioReal = 82.5
        Case 214                  '2 mediano clasico x 179
        PrecioReal = 89.5
        Case 236                  '2 mediano clasico x 179
        PrecioReal = 90
        Case 146                  '2 waffle especiales x 219
        PrecioReal = 109.5
        Case 195                  '2 waffle especiales x 219
        PrecioReal = 120
        Case 218                  '2 waffle especiales x 275
        PrecioReal = 137.5
        Case 120                   'Agranda batido por 5 cordobas de 16oz a 20oz
        PrecioReal = mont - 11
        Case 131                   'Agranda batido por 5 cordobas de 16oz a 20oz
        PrecioReal = mont - 6
        Case 140                   'Agranda batido por 5 cordobas de 16oz a 20oz
        PrecioReal = mont - 9
        Case 220                   'Agranda batido por 5 cordobas de 16oz a 20oz
            If rst("CodGrupo") = 2 Or rst("CodGrupo") = 1 Or rst("CodGrupo") = 24 Then 'premium, eseciales, con proteina
                PrecioReal = mont - 5
            ElseIf rst("CodGrupo") = 8 Then 'saludables suoer vedes
                PrecioReal = mont
            ElseIf rst("CodGrupo") = 3 Then 'clasico
                PrecioReal = mont - 10
            ElseIf rst("CodGrupo") = 16 Then 'limonada
                If rst("CodBatido") Like "L004*" Then 'limonada de fresa
                    PrecioReal = mont - 5
                Else 'resto de limonadas
                    PrecioReal = mont - 10
                End If
            Else
                PrecioReal = mont - 5
            End If
        Case 234                   'Agranda batido por 5 cordobas de 16oz a 20oz
            If rst("CodGrupo") = 2 Or rst("CodGrupo") = 1 Or rst("CodGrupo") = 24 Then 'premium, eseciales, con proteina
                PrecioReal = mont - 10
            ElseIf rst("CodGrupo") = 8 Then 'saludables suoer vedes
                PrecioReal = mont - 10
            ElseIf rst("CodGrupo") = 3 Then 'clasico
                PrecioReal = mont - 10
            ElseIf rst("CodGrupo") = 16 Then 'limonada
                PrecioReal = mont - 10
            Else
                PrecioReal = mont - 10
            End If
        Case 191                   'Agranda batido por 7 cordobas de 16oz a 20oz
        PrecioReal = mont - 8
        Case 147                   'Agranda batido por 3 puntos de 16oz a 20oz
        PrecioReal = mont - 14
        Case 221                  'Agranda batido por 3 puntos de 16oz a 20oz
                If rst("CodGrupo") = 2 Or rst("CodGrupo") = 1 Or rst("CodGrupo") = 8 Or rst("CodGrupo") = 4 Then 'premium, eseciales, con proteina
                    PrecioReal = mont - 10
                ElseIf rst("CodGrupo") = 3 Then 'saludables suoer vedes
                    PrecioReal = mont - 15
                ElseIf rst("CodGrupo") = 16 Then 'limonada
                    If rst("CodBatido") Like "L004*" Then 'limonada de fresa
                        PrecioReal = mont - 10
                    Else 'resto de limonadas
                        PrecioReal = mont - 15
                    End If
                Else
                    PrecioReal = mont - 10
                End If
        Case 235                  'Agranda batido por 3 puntos de 16oz a 20oz
            If rst("CodGrupo") = 2 Or rst("CodGrupo") = 1 Or rst("CodGrupo") = 8 Or rst("CodGrupo") = 4 Then 'premium, eseciales, con proteina
                PrecioReal = mont - 15
            ElseIf rst("CodGrupo") = 3 Then 'saludables suoer vedes
                PrecioReal = mont - 15
            ElseIf rst("CodGrupo") = 16 Then 'limonada
                PrecioReal = mont - 15
            Else
                PrecioReal = mont - 15
            End If
        Case 122                   'Canjeado con cupon d decuento de 8 cordobas
        PrecioReal = mont - 8
        Case 123                   'Canjeado con cupon d decuento de 9 cordobas
        PrecioReal = mont - 9
        Case 124                   'Canjeado con cupon d decuento de 10 cordobas
        PrecioReal = mont - 10
        
        Case 171                    ' precio unico de feria
        PrecioReal = 130
        Case 211                    ' precio unico de feria con 50% descuento
        PrecioReal = 65
        Case 200                    ' precio unico de feria
        PrecioReal = 120
        Case 176                    ' precio unico de feria
        PrecioReal = 140
        Case 177                    ' precio unico de feria
        PrecioReal = 80
        Case 178                    ' precio unico de feria
        PrecioReal = 150
        Case 180                 ' precio unico de feria
        PrecioReal = 30
        Case 201                 ' precio unico de feria
        PrecioReal = 40
        Case 181                    ' precio unico de feria
        PrecioReal = 120
        Case 182                    ' precio unico de feria
        PrecioReal = 40
        Case 183                    ' precio unico de feria
        PrecioReal = 99
        
        Case 130                   'Agranda batido por compra de semilla de 16 a 20oz
        PrecioReal = mont - 11
        Case 149                   'Agranda batido por compra de semilla de 16 a 20oz
        PrecioReal = mont - 14
        Case 163                   'Agranda batido por compra de semilla de 16 a 20oz
        PrecioReal = mont - 14
        Case 164                   'Agranda batido por compra de semilla de 16 a 20oz
        PrecioReal = mont - 9
        Case 112                ' 10 pesos menos si viene x influencer
            If rst("CodGrupo") = 2 Then
                PrecioReal = mont - 16 + 5
            Else
                PrecioReal = mont - 15 + 5
            End If
        Case 36                    ' nuevos productos pan y paz
            If rst("Medida") = "Mediano" Then
                If rst("CodGrupo") = 2 Then
                    PrecioReal = 63
                ElseIf rst("CodGrupo") = 8 Then
                    PrecioReal = 54
                ElseIf rst("CodGrupo") = 1 Then
                    PrecioReal = 54
                ElseIf rst("CodGrupo") = 3 Then
                    PrecioReal = 45
                End If
            ElseIf rst("Medida") = "Gigantona" Then
                If rst("CodGrupo") = 2 Then
                    PrecioReal = 70
                ElseIf rst("CodGrupo") = 8 Then
                    PrecioReal = 61
                ElseIf rst("CodGrupo") = 1 Then
                    PrecioReal = 61
                ElseIf rst("CodGrupo") = 3 Then
                    PrecioReal = 52
                End If
            End If
        
        Case Else
        PrecioReal = mont
    End Select
    
Else
    
    PrecioReal = 0 'pedido anulado

End If

rst.Close
     
Exit Function

Nulo:
PrecioReal = 0
End Function


Function IngresarPedido(cbatido As String, cpedido As Long, vincu As Long, Optional promox As Integer = 5, Optional escombo As Integer = 0, Optional esgrupocombo As String = "", Optional cantix As Integer = 1)

'On Error GoTo Again
Dim Cant As Integer
Dim grupi As Integer
Dim obse As String
Dim nuevaobse As String
Dim subpe As Long
Dim formName As String

Dim tipovinculo As Integer

Dim rst As DAO.Recordset
Dim miSQL As String

grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & cbatido & "'")
tipovinculo = IIf(vincu = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & vincu & "") & "'"))

If cpedido <> 0 Then ' Se abrio sin tener nota de pedido abierto
    
    If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
        If IsNull([Forms]![Nota de Pedido]![HoraIngresoProducto]) Then
            [Forms]![Nota de Pedido]![HoraIngresoProducto] = Time
        End If
    End If
    
    Select Case grupi
        Case 20
        'Combos elegir
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
            " VALUES ('" & cbatido & "',1, " & cpedido & ", " & promox & ", 0, '-', " & vincu & ")"
            DoCmd.SetWarnings True
           
            subpe = UltimoSubPedido()
            
            'Recorrer todas las notas del pedido del dia
            miSQL = "SELECT CombosComponentes.CodBatido, CombosComponentes.Cantidad, CombosComponentes.Menu, CombosComponentes.GrupoComponente, CombosComponentes.ProductoEspecifico FROM CombosComponentes" & _
            " WHERE (((CombosComponentes.CodBatido)='" & cbatido & "'))"
            Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
            rst.MoveFirst
            Do While Not rst.EOF
                
                If rst("GrupoComponente") = 0 Then
                    Call IngresarPedidoDirectoUnidad(rst("ProductoEspecifico"), cpedido, subpe, 104, , rst("Cantidad"))
                Else
                    formName = rst("Menu")
                    DoCmd.OpenForm formName
                    Forms(formName)!vinculos = subpe
                    Forms(formName)!grupocombo = formName
                    Forms(formName)!cantidadtotal = rst("Cantidad")
                End If
        
                rst.MoveNext
            Loop
    
        Case 1, 2, 3, 4, 8, 16, 18, 19, 24
        'productos con endulzantes
            DoCmd.OpenForm "Menu Endulzantes"
            [Forms]![Menu Endulzantes]![cpedido] = cpedido
            [Forms]![Menu Endulzantes]![nbatido] = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & cbatido & "'") & " " & DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & cbatido & "'")
            
            [Forms]![Menu Endulzantes]![ccantidad] = cantix
            [Forms]![Menu Endulzantes]![cbatido] = cbatido
            [Forms]![Menu Endulzantes]![vinculos] = vincu
            [Forms]![Menu Endulzantes]![menucombo] = esgrupocombo
            If escombo = 1 Then
                [Forms]![Menu Endulzantes]![ccantidad].Enabled = False
            End If
        
        Case Else
            If CurrentProject.AllForms("VentasCodigoBarras").IsLoaded Or tipovinculo = 20 Then
                Cant = 1
            Else
                If escombo = 1 Then
                    Cant = cantix
                Else
                    Cant = InputBox("Ingrese Cantidad", "CANTIDAD DE PEDIDO", "1", 7000, 5000)
                End If
            End If
            
            Dim promoActual As Integer
            promoActual = IIf(tipovinculo = 20, 104, promox)
            
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
            " VALUES ('" & cbatido & "', " & Cant & ", " & cpedido & ", " & promoActual & ", 0, '-', " & vincu & ")"
            DoCmd.SetWarnings True
            
            If grupi = 11 And CurrentProject.AllForms("Menu Adicionales").IsLoaded Then
            'Cuando se agrega un producto adicional , se anexa el detalle al producto principal, esta en menu de adicionales
            
                obse = DLookup("[Observaciones]", "[SubPedido]", "[CodSubPedido]=" & [Forms]![Menu Adicionales]![asubpedido])
                
                nuevaobse = obse & ", " & fraccion(CDbl(Cant)) & " extra de " & nombreproductoventa(cbatido)
                
                DoCmd.SetWarnings False
                DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.Observaciones = '" & nuevaobse & "'" & _
                " WHERE SubPedido.CodSubPedido = " & [Forms]![Menu Adicionales]![asubpedido]
                DoCmd.SetWarnings True
            
            End If
            
            'cerrar ventana de producto especifico
            If tipovinculo = 20 Then
                DoCmd.Close acForm, esgrupocombo
            End If

            
    End Select

End If

Exit Function
Again:
MsgBox "Se produjo un error al agregar el pedido, ingrese nuevamente la cantidad"
End Function

Function IngresarPedidoDirectoUnidad(cbatido As String, cpedido As Long, vincu As Long, Optional promox As Integer = 5, Optional escombo As Integer = 0, Optional cantix As Integer = 1)

On Error GoTo Again
Dim Cant As Integer
Dim obse As String
Dim nuevaobse As String
Dim subpe As Long

If cpedido <> 0 Then ' Se abrio sin tener nota de pedido abierto
    
    If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
        If IsNull([Forms]![Nota de Pedido]![HoraIngresoProducto]) Then
            [Forms]![Nota de Pedido]![HoraIngresoProducto] = Time
        End If
    End If
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & cbatido & "', " & cantix & ", " & cpedido & ", " & promox & ", 0, '-', " & vincu & ")"
    DoCmd.SetWarnings True
    
    MsgBox "Producto agregado correctamente"
        
End If

Exit Function
Again:
MsgBox "Se produjo un error al agregar el pedido, ingrese nuevamente la cantidad"
End Function
Function IngresarPedidoParaCombo(cbatido As String, cpedido As Double, comb As String, rut As Integer)

On Error GoTo Again
Dim nombrebat As String
nombrebat = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & comb & "'")
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, SinAzucar, Azucar)" & _
" VALUES ('" & cbatido & "',1, " & cpedido & ", 104, '" & nombrebat & "', 0, '-')"
DoCmd.SetWarnings True

Select Case comb
    Case "combo1"
        Select Case rut
            Case 1
            DoCmd.OpenForm "MenuParaComboPremium"
            [Forms]![MenuParaComboPremium]![combodec] = comb
            [Forms]![MenuParaComboPremium]![Etiqueta2059].Caption = "ELEGIR BATIDO PREMIUM"
            [Forms]![MenuParaComboPremium]![ruta] = rut + 1
            DoCmd.Close acForm, "MenuParaComboWaffles"
            Case 2
            DoCmd.Close acForm, "MenuParaComboPremium"
        End Select
    Case "combo2"
        Select Case rut
            Case 1, 2, 3
            'DoCmd.OpenForm "MenuParaComboClasicos"
            [Forms]![MenuParaComboClasicos]![combodec] = comb
            [Forms]![MenuParaComboClasicos]![Etiqueta2059].Caption = "ELEGIR BATIDO #" & rut + 1 & " CLASICO"
            [Forms]![MenuParaComboClasicos]![ruta] = rut + 1
            'DoCmd.Close acForm, "MenuParaComboWaffles"
            Case 4
            DoCmd.Close acForm, "MenuParaComboClasicos"
        End Select
    Case "combo3"
        Select Case rut
            Case 1
            DoCmd.OpenForm "MenuParaComboClasicos"
            [Forms]![MenuParaComboClasicos]![combodec] = comb
            [Forms]![MenuParaComboClasicos]![Etiqueta2059].Caption = "ELEGIR BATIDO CLASICO"
            [Forms]![MenuParaComboClasicos]![ruta] = rut + 1
            DoCmd.Close acForm, "MenuParaComboWaffles"
            Case 2
            DoCmd.Close acForm, "MenuParaComboClasicos"
        End Select
    Case "combo4"
        Select Case rut
            Case 1
            'DoCmd.OpenForm "MenuParaComboSaludables"
            [Forms]![MenuParaComboSaludables]![combodec] = comb
            [Forms]![MenuParaComboSaludables]![Etiqueta2059].Caption = "ELEGIR BATIDO #" & rut + 1 & " SALUDABLE"
            [Forms]![MenuParaComboSaludables]![ruta] = rut + 1
            'DoCmd.Close acForm, "MenuParaComboSaludables"
            Case 2
            DoCmd.Close acForm, "MenuParaComboSaludables"
        End Select
    Case "combo5"
        Select Case rut
            Case 1
            'DoCmd.OpenForm "MenuParaComboWaffles"
            [Forms]![MenuParaComboWaffles]![combodec] = comb
            [Forms]![MenuParaComboWaffles]![Etiqueta2059].Caption = "ELEGIR WAFFLE #" & rut + 1
            [Forms]![MenuParaComboWaffles]![ruta] = rut + 1
            'DoCmd.Close acForm, "MenuParaComboSaludables"
            Case 2
            DoCmd.Close acForm, "MenuParaComboWaffles"
        End Select
    Case Else
        MsgBox "No existe combo registrado"
End Select

Exit Function
Again:
MsgBox "Se produjo un error al agregar el pedido, ingrese nuevamente la cantidad"
End Function



Function ElegirTipoEnvase() As Long

If MsgBox("PARA CONSUMIR EN MESA (VIDRIO)?", vbYesNo, "PREFERENCIA DE VASO") = vbYes Then
ElegirTipoEnvase = -1
[Forms]![Nota de Pedido]![bvidrio].Caption = "VIDRIO"
Else
ElegirTipoEnvase = 0
[Forms]![Nota de Pedido]![bvidrio].Caption = "PLASTICO"
End If

End Function

Function IngresarNombre() As String

Dim Nombre As String

Nombre = StrConv(InputBox("Ingresar Nombre de Cliente", "Nombre de Cliente"), vbProperCase)

If Len(Nombre) < 2 Then
MsgBox "INGRESAR NOMBRE CORRECTAMENTE"
Nombre = IngresarNombre()
End If

IngresarNombre = Nombre

End Function

Function IngresarClub() As Long

On Error GoTo NoClub
Dim rst As DAO.Recordset
Dim miSQL As String

If MsgBox("Cuenta con Membresia CLUB PITAYA?", vbYesNo, "CLUB PITAYA") = vbYes Then

IngresarClub = CInt(InputBox("Ingresar Codigo de CLUB PITAYA", "CLIENTE CLUB PITAYA"))

'Buscar Codigo CLUB
miSQL = "SELECT ClientesClub.CodCliente, ClientesClub.Nombre FROM ClientesClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb' WHERE (ClientesClub.CodCliente=" & IngresarClub & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MsgBox "BIENVENID@ " & rst("Nombre") & " "
rst.Close

Else
IngresarClub = 0

End If

Exit Function

NoClub:
MsgBox "Codigo Erroneo"
IngresarClub = IngresarClub()

End Function


Function agradecimiento(Tipo As Integer) As String

If Tipo <> 0 Then

agradecimiento = "Gracias por tu apoyo en cuidar el medio ambiente y utilizar envase de vidrio"

Else

agradecimiento = "Ayudanos a cuidar el medio ambiente utilizando envase de vidrio"
End If


End Function

Function DefinirPrimeraVez() As String

If MsgBox("Primera Vez en PITAYA?", vbYesNo, "PRIMERA VEZ") = vbYes Then
DefinirPrimeraVez = -1

Else
DefinirPrimeraVez = 0

End If

End Function

Function MontoPedido(codigo As Long) As Double

'Ingresos brutos totales de cada pedido (nota de pedido)

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la suma total de la lista de pedido de nota de pedido
miSQL = "SELECT NotaDePedido.CodPedido, Sum(PrecioReal([SubPedido]![CodSubPedido])*[SubPedido]![Cantidad]*(1+[NotaDepedido]![Propina])) AS total" & _
" FROM NotaDePedido INNER JOIN (DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido)" & _
" ON NotaDePedido.CodPedido = SubPedido.CodPedido GROUP BY NotaDePedido.CodPedido" & _
" HAVING (((NotaDePedido.CodPedido)=" & codigo & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MontoPedido = rst("total")
rst.Close
     
Exit Function

Nulo:
MontoPedido = 0

End Function

Function nombreactualmedidas(medi As String) As String
On errror GoTo Nulo
Select Case medi
    Case "Mediano"
        nombreactualmedidas = "Pequeño"
    Case "Gigantona"
        nombreactualmedidas = "Normal"
    Case "Bowl"
        nombreactualmedidas = "Bowl"
    Case "Club Pitaya"
        nombreactualmedidas = "Unid"
    Case "Kid"
        nombreactualmedidas = "Kid"
    Case "MiniBowl"
        nombreactualmedidas = "MiniBowl"
    Case "NV"
        nombreactualmedidas = "Unid"
    Case Else
        nombreactualmedidas = "Unid"
End Select
Exit Function

Nulo:
nombreactualmedidas = "Unid"
End Function
Function VentaRealSemanal(sem As Integer) As Long

'Ingresos brutos totales de una semana especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de monto total de una semana con monto de pedido
miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Semana, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto FROM NotaDePedido GROUP BY numerosemana([NotaDePedido]![Fecha]) HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentaRealSemanal = rst("Monto")
rst.Close
     
Exit Function

Nulo:
VentaRealSemanal = 0

End Function

Function VentaTeoricoSemanal(sem As Integer) As Long

'Precio de venta de batidos totales de una semana especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de precios reales vendidos de productos
miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Semana, Sum([DBBatidos]![Precio]*[SubPedido]![Cantidad]) AS Monto, NotaDePedido.Anulado" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), NotaDePedido.Anulado" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sem & ") AND ((NotaDePedido.Anulado)=0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentaTeoricoSemanal = rst("Monto")
rst.Close
     
Exit Function

Nulo:
VentaTeoricoSemanal = 0

End Function



Function confirmarclub(club As Long) As Long

On Error GoTo NoClub
Dim rst As DAO.Recordset
Dim miSQL As String

'Buscar Codigo CLUB en Base de datos
miSQL = "SELECT ClientesClub.CodCliente, ClientesClub.Nombre FROM ClientesClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb' WHERE (ClientesClub.CodCliente=" & club & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
confirmarclub = rst("CodCliente")
rst.Close

Exit Function

NoClub:
confirmarclub = 0

End Function

Function confirmarclub2(club As Long) As Long
On Error GoTo Nulo

Dim rst As DAO.Recordset
Dim miSQL As String

'Buscar Codigo CLUB en Base de datos
miSQL = "SELECT ClientesClub.CodCliente, ClientesClub.Nombre FROM ClientesClub" & _
" WHERE (ClientesClub.CodCliente=" & club & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
confirmarclub2 = rst("CodCliente")
rst.Close

Exit Function

Nulo:
confirmarclub2 = 0

End Function

Function montoposdia(Fecha As Date) As Long
'Monto de total facturado con POS de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia
miSQL = "SELECT NotaDePedido.Anulado, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto, NotaDePedido.POS, NotaDePedido.Fecha " & _
"FROM NotaDePedido GROUP BY NotaDePedido.Anulado, NotaDePedido.POS, NotaDePedido.Fecha " & _
"HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.POS)<>0) AND ((NotaDePedido.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montoposdia = rst("Monto")
rst.Close

Exit Function

NoPos:
montoposdia = 0

End Function

Function montoposdiaguardado(Fecha As Date) As Long
'Monto de total facturado con POS de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia
miSQL = "SELECT NotaDePedido.Anulado, Sum(NotaDePedido.TotalGuardado) AS Monto, NotaDePedido.POS, NotaDePedido.Fecha " & _
"FROM NotaDePedido GROUP BY NotaDePedido.Anulado, NotaDePedido.POS, NotaDePedido.Fecha " & _
"HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.POS)<>0) AND ((NotaDePedido.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montoposdiaguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montoposdiaguardado = 0

End Function

Function montoposdiaguardadoTransferencia(fech As Date) As Long
'Monto de total facturado con POS de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia

miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision," & _
" NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & fech & "#)" & _
" AND ((NotaDePedido.POS)<>0) AND ((Delivery.Comision)=0) AND ((NotaDePedido.Transferencia)<>0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montoposdiaguardadoTransferencia = rst("Monto")
rst.Close

Exit Function

NoPos:
montoposdiaguardadoTransferencia =  0

End Function

Function montoposdiahastahoraguardadoTransferencia(fech As Date, horax As Date) As Long
'Monto de total facturado con POS de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia

miSQL = "SELECT NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & horax & "# AS horalimite," & _
" NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & horax & "#, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Fecha)=#" & fech & "#) AND (([NotaDePedido]![Hora]<#" & horax & "#)<>0)" & _
" AND ((NotaDePedido.POS)<>0) AND ((Delivery.Comision)=0) AND ((NotaDePedido.Transferencia)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montoposdiahastahoraguardadoTransferencia = rst("Monto")
rst.Close

Exit Function

NoPos:
montoposdiahastahoraguardadoTransferencia =  0

End Function

Function montohugodia(fech As Date) As Long
'Monto de total facturado con HUGO de un dia especifico de pedidos no anulados, tambien pedidosya y piki

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia
miSQL = "SELECT NotaDePedido.Anulado, [NotaDePedido]![Delivery]=5 Or [NotaDePedido]![Delivery]=6 Or [NotaDePedido]![Delivery]=8 AS CondicionDelivery," & _
" NotaDePedido.Fecha, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto" & _
" FROM NotaDePedido " & _
" GROUP BY NotaDePedido.Anulado, [NotaDePedido]![Delivery]=5 Or [NotaDePedido]![Delivery]=6 Or [NotaDePedido]![Delivery]=8, NotaDePedido.Fecha" & _
" HAVING (((NotaDePedido.Anulado)=0) AND (([NotaDePedido]![Delivery]=5 Or [NotaDePedido]![Delivery]=6 Or [NotaDePedido]![Delivery]=8)<>0) AND ((NotaDePedido.Fecha)=#" & fech & "#));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montohugodia = rst("Monto")
rst.Close

Exit Function

NoPos:
montohugodia = 0

End Function

Function montohugodiaguardado(fech As Date) As Long
'Monto de total facturado con HUGO de un dia especifico de pedidos no anulados, tambien pedidosya y piki

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia

miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision," & _
" NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & fech & "#) AND ((NotaDePedido.POS)<>0) AND ((Delivery.Comision)<>0) AND ((NotaDePedido.Transferencia)=0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montohugodiaguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montohugodiaguardado =  0

End Function

Function montohugodiahastahoraguardado(fech As Date, horax As Date) As Long
'Monto de total facturado con HUGO de un dia especifico de pedidos no anulados, tambien pedidosya y piki

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia

miSQL = "SELECT NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & horax & "# AS horalimite," & _
" NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & horax & "#, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Fecha)=#" & fech & "#) AND (([NotaDePedido]![Hora]<#" & horax & "#)<>0)" & _
" AND ((NotaDePedido.POS)<>0) AND ((Delivery.Comision)<>0) AND ((NotaDePedido.Transferencia)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montohugodiahastahoraguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montohugodiahastahoraguardado =  0

End Function

Function montosoloposdiaguardado(fech As Date) As Long
'Monto de total facturado con POS solamente de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia
miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS," & _
" Delivery.Comision, NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & fech & "#)" & _
" AND ((NotaDePedido.POS)<>0) AND ((Delivery.Comision)=0) AND ((NotaDePedido.Transferencia)=0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montosoloposdiaguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montosoloposdiaguardado =  0

End Function

Function montosoloposdiahastahoraguardado(fech As Date, horax As Date) As Long
'Monto de total facturado con POS solamente de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia
miSQL = "SELECT NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & horax & "# AS horalimite," & _
" NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Fecha, [NotaDePedido]![Hora]<#" & horax & "#, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Fecha)=#" & fech & "#) AND (([NotaDePedido]![Hora]<#" & horax & "#)<>0)" & _
" AND ((NotaDePedido.POS)<>0) AND ((Delivery.Comision)=0) AND ((NotaDePedido.Transferencia)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montosoloposdiahastahoraguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montosoloposdiahastahoraguardado =  0

End Function

Function montosoloefectivodiaguardado(fech As Date) As Long
'Monto de total facturado con efectivo solamente de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de efectivo del dia
miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS," & _
" Delivery.Comision, NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & fech & "#)" & _
" AND ((NotaDePedido.POS)=0) AND ((Delivery.Comision)=0) AND ((NotaDePedido.Transferencia)=0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montosoloefectivodiaguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montosoloefectivodiaguardado =  0

End Function

Function ventaproductosemana(prod As String, Med As String, sem As Integer) As Long
'Cantidad vendida de un producto y medda especifico en una semana especifica

On Error GoTo NoVentas
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de cantidad suma de productos vvendido en semana especifica
miSQL = "SELECT DBBatidos.Nombre, DBBatidos.Medida, Sum(SubPedido.Cantidad) AS SumaDeCantidad, numerosemana([NotaDePedido]![Fecha]) AS semana, NotaDePedido.Anulado" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY DBBatidos.Nombre, DBBatidos.Medida, numerosemana([NotaDePedido]![Fecha]), NotaDePedido.Anulado" & _
" HAVING (DBBatidos.Nombre='" & prod & "' AND DBBatidos.Medida='" & Med & "' AND numerosemana([NotaDePedido]![Fecha])=" & sem & " AND NotaDePedido.Anulado=0)"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ventaproductosemana = rst("SumaDeCantidad")
rst.Close

Exit Function

NoVentas:
ventaproductosemana = 0

End Function

Function ventaMaximoproductosemana(prod As String, Med As String, sem As Integer, rango As Integer) As Long
'Cantidad vendida de un producto y medda especifico en una semana especifica

On Error GoTo NoVentas
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de cantidad suma de productos vvendido en semana especifica
miSQL = "SELECT NotaDePedido.Anulado, numerosemana([NotaDePedido]![Fecha]) AS semana, DBBatidos.Nombre," & _
" DBBatidos.Medida, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (NotaDePedido" & _
" INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos" & _
" ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, numerosemana([NotaDePedido]![Fecha]), DBBatidos.Nombre, DBBatidos.Medida" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((numerosemana([NotaDePedido]![Fecha])) Between " & sem - rango & " And " & sem & ")" & _
" AND ((DBBatidos.Nombre)='" & prod & "') AND ((DBBatidos.Medida)='" & Med & "'))" & _
" ORDER BY Sum(SubPedido.Cantidad) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ventaMaximoproductosemana = rst("SumaDeCantidad")
rst.Close

Exit Function

NoVentas:
ventaMaximoproductosemana = 0

End Function

Function ventaMaximoproductosemanalocal(prod As String, Med As String, sem As Integer, rango As Integer, locy As Integer) As Long
'Cantidad vendida de un producto y medda especifico en una semana especifica

On Error GoTo NoVentas
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de cantidad suma de productos vvendido en semana especifica
miSQL = "SELECT NotaDePedido" & locy & ".Anulado, numerosemana([NotaDePedido" & locy & "]![Fecha]) AS semana, DBBatidos.Nombre," & _
" DBBatidos.Medida, Sum(SubPedido" & locy & ".Cantidad) AS SumaDeCantidad" & _
" FROM DBBatidos INNER JOIN (SubPedido" & locy & " INNER JOIN NotaDePedido" & locy & "" & _
" ON SubPedido" & locy & ".CodPedido = NotaDePedido" & locy & ".CodPedido) ON DBBatidos.CodBatido = SubPedido" & locy & ".CodBatido" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY NotaDePedido" & locy & ".Anulado, numerosemana([NotaDePedido" & locy & "]![Fecha]), DBBatidos.Nombre, DBBatidos.Medida" & _
" HAVING (((NotaDePedido" & locy & ".Anulado)=0) AND ((numerosemana([NotaDePedido" & locy & "]![Fecha])) Between " & sem - rango & " And " & sem & ")" & _
" AND ((DBBatidos.Nombre)='" & prod & "') AND ((DBBatidos.Medida)='" & Med & "'))" & _
" ORDER BY Sum(SubPedido" & locy & ".Cantidad) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ventaMaximoproductosemanalocal = rst("SumaDeCantidad")
rst.Close

Exit Function

NoVentas:
ventaMaximoproductosemanalocal = 0

End Function

Function ConsumoProductoMostradorSemanaLocal(prod As String, Med As String, sem As Integer, locy As Integer) As Long
'Cantidad vendida de un producto y medda especifico en una semana especifica

On Error GoTo NoVentas
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de cantidad suma de productos vvendido en semana especifica
miSQL = "SELECT NotaDePedido" & locy & ".Anulado, numerosemana([NotaDePedido" & locy & "]![Fecha]) AS semana, DBBatidos.Nombre," & _
" DBBatidos.Medida, Sum(SubPedido" & locy & ".Cantidad) AS SumaDeCantidad" & _
" FROM DBBatidos INNER JOIN (SubPedido" & locy & " INNER JOIN NotaDePedido" & locy & "" & _
" ON SubPedido" & locy & ".CodPedido = NotaDePedido" & locy & ".CodPedido) ON DBBatidos.CodBatido = SubPedido" & locy & ".CodBatido" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY NotaDePedido" & locy & ".Anulado, numerosemana([NotaDePedido" & locy & "]![Fecha]), DBBatidos.Nombre, DBBatidos.Medida" & _
" HAVING (((NotaDePedido" & locy & ".Anulado)=0) AND (numerosemana([NotaDePedido" & locy & "]![Fecha])=" & sem & ")" & _
" AND ((DBBatidos.Nombre)='" & prod & "') AND ((DBBatidos.Medida)='" & Med & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoProductoMostradorSemanaLocal = rst("SumaDeCantidad")
rst.Close

Exit Function

NoVentas:
ConsumoProductoMostradorSemanaLocal = 0

End Function

Function OperarioCaja(hor As Date, fech As Date) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CierreDiario.Fecha, CierreDiario.HoraInicial, CierreDiario.HoraFinal, CierreDiario.CodOperario" & _
" FROM CierreDiario INNER JOIN Operarios ON CierreDiario.CodOperario = Operarios.CodOperario" & _
" WHERE (((CierreDiario.Fecha)=#" & fech & "#) AND ((CierreDiario.HoraInicial)<=#" & hor & "#) AND ((CierreDiario.HoraFinal)>=#" & hor & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
OperarioCaja = NombreOperario(rst("CodOperario"))
rst.Close

Exit Function

Nulo:
OperarioCaja = "No Registrado"

End Function


Function UltimoSubPedido() As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TOP 1 SubPedido.CodSubPedido" & _
" FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" ORDER BY SubPedido.CodSubPedido DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
UltimoSubPedido = rst("CodSubPedido")
rst.Close

Exit Function

Nulo:
UltimoSubPedido = UltimoSubPedido()

End Function

Function LlevaEndulzante(codbat As String) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodIngrediente, SubReceta.CodBatido, SubReceta.Cantidad" & _
" FROM SubReceta" & _
" WHERE (((SubReceta.CodIngrediente)='E004') AND ((SubReceta.CodBatido)='" & codbat & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
If rst("Cantidad") > 0 Then
    LlevaEndulzante = 1
Else
    LlevaEndulzante = 0
End If
rst.Close

Exit Function

Nulo:
LlevaEndulzante = 0

End Function

Function DescuentoDelivery(coddelivery As Integer) As Double
If coddelivery = 0 Or coddelivery = 5 Or coddelivery = 6 Or coddelivery = 7 Or coddelivery = 8 Then DescuentoDelivery = 0: Exit Function
DescuentoDelivery = Choose(coddelivery, 5, 10, 10, 10)
End Function

Function CantidadxTipoDelivery(fech As Date, tip As Integer) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Fecha, NotaDePedido.Delivery, Sum(1) AS Total" & _
" FROM NotaDePedido GROUP BY NotaDePedido.Fecha, NotaDePedido.Delivery" & _
" HAVING (NotaDePedido.Fecha=#" & fech & "# AND NotaDePedido.Delivery=" & tip & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadxTipoDelivery = rst("Total")
rst.Close
Exit Function

Nulo:
CantidadxTipoDelivery = 0
End Function

Function CantidadProductosxPromocionDia(fech As Date, promoc As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.Fecha, NotaDePedido.Anulado, SubPedido.CodPromocion, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido" & _
" GROUP BY NotaDePedido.Fecha, NotaDePedido.Anulado, SubPedido.CodPromocion" & _
" HAVING (((NotaDePedido.Fecha)=#" & fech & "#) AND ((NotaDePedido.Anulado)=0) AND ((SubPedido.CodPromocion)=" & promoc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadProductosxPromocionDia = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
CantidadProductosxPromocionDia = 0
End Function

Function ultimanotapedido() As Long
'ultimo codigo de nota de pedido
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodPedido FROM NotaDePedido ORDER BY CodPedido DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimanotapedido = rst("CodPedido")
rst.Close

Exit Function

Nulo:
ultimanotapedido = 0

End Function


Function AdicionalVigenteDeIngrediente(ainge As String) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.Precio, SubReceta.CodIngrediente, DBBatidos.CodGrupo, DBBatidos.Vigencia, DBBatidos.CodBatido" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" WHERE (((SubReceta.CodIngrediente)='" & ainge & "') AND ((DBBatidos.CodGrupo)=11) AND ((DBBatidos.Vigencia)<>0))" & _
" ORDER BY DBBatidos.Precio DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AdicionalVigenteDeIngrediente = rst("CodBatido")
rst.Close
Exit Function

Nulo:
AdicionalVigenteDeIngrediente = ""

End Function
Function CrearNotaDePedido(tip As Integer) As String
If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
    MsgBox "Ya hay un pedido abierto, primero cerrar el pedido para abrir uno nuevo"
    Exit Function
End If


' tip
    ''1 normal efectivo
    ''2 Hugo
    ''3 PedidosYa
    ''4 Con delivery motorizado
    
'''''''Version crear primero lueg abrir
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO NotaDePedido(Fecha, HoraCreado)" & _
" values (#" & Date & "#, #" & Time & "#)"
'DoCmd.RunSQL "INSERT INTO NotaDePedido(Fecha)" & _
'" values (#" & Date & "#)"
DoCmd.SetWarnings True

Dim codnota As Long
codnota = ultimanotapedido()
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & codnota

[Forms]![Nota de Pedido]![Texto166] = 0

If CurrentProject.AllForms("Main Pitaya").IsLoaded Then
    [Forms]![Nota de Pedido]![colaborador] = [Forms]![Main Pitaya]![codigologin]
End If

[Forms]![Nota de Pedido].Form.Requery

Select Case tip
    Case 1
        [Forms]![Nota de Pedido]![tipofactura] = "PEDIDO EN TIENDA"
        [Forms]![Nota de Pedido]![pacumulado] = 0
        [Forms]![Nota de Pedido]![TipoDelivery] = 0
        [Forms]![Nota de Pedido]![Delivery] = 0
        [Forms]![Nota de Pedido]![POS] = 0
        [Forms]![Nota de Pedido]![Transferencia] = 0
        [Forms]![Nota de Pedido]![bpos].Caption = "EFECTIVO"
        [Forms]![Nota de Pedido]![bpos].Enabled = True
        [Forms]![Nota de Pedido]![Comando68].Visible = True '  menu de clientes normal\
        [Forms]![Nota de Pedido]![Comando1204].Visible = False ' ocultar menu de pedidoya
        '[Forms]![Nota de Pedido]![Etiqueta1386].Visible = False ' rotulo de eleccion de motorizado
        '[Forms]![Nota de Pedido]![CodMotorizado].Visible = False ' lista de eleccion de motorizado
        
        [Forms]![Nota de Pedido]![Comando1669].Visible = False ' Mostrar bton de datos de envio
        [Forms]![Nota de Pedido]![botonaprobar].Visible = False ' boton de aprobar
        
        DoCmd.OpenForm "Menu PITAYA Global"
    Case 2
        [Forms]![Nota de Pedido]![tipofactura] = "PEDIDO POR APLICACION HUGO"
        [Forms]![Nota de Pedido]![pacumulado] = 0
        [Forms]![Nota de Pedido]![TipoDelivery] = 5
        [Forms]![Nota de Pedido]![Delivery] = 5
        [Forms]![Nota de Pedido]![POS] = -1
        [Forms]![Nota de Pedido]![Transferencia] = 0
        [Forms]![Nota de Pedido]![bpos].Caption = "HUGO"
        [Forms]![Nota de Pedido]![bpos].Enabled = False
        [Forms]![Nota de Pedido]![Comando68].Visible = False ' Ocultar menu de clientes normal
        [Forms]![Nota de Pedido]![Comando1204].Visible = False ' mostrar menu de pedidosya normal\
        [Forms]![Nota de Pedido]![Etiqueta1386].Visible = False ' rotulo de eleccion de motorizado
        [Forms]![Nota de Pedido]![CodMotorizado].Visible = False ' lista de eleccion de motorizado
        
        [Forms]![Nota de Pedido]![Comando1669].Visible = False ' Mostrar bton de datos de envio
        [Forms]![Nota de Pedido]![botonaprobar].Visible = False ' boton de aprobar
        
        [Forms]![Nota de Pedido]![tipopropina].Visible = False
        '[Forms]![Nota de Pedido]![CodCliente].Enabled = False
        DoCmd.OpenForm "Menu PITAYA Delivery"
    Case 3
        [Forms]![Nota de Pedido]![tipofactura] = "PEDIDOS YA"
        [Forms]![Nota de Pedido]![pacumulado] = 0
        [Forms]![Nota de Pedido]![TipoDelivery] = 8
        [Forms]![Nota de Pedido]![Delivery] = 8
        [Forms]![Nota de Pedido]![POS] = -1
        [Forms]![Nota de Pedido]![Transferencia] = 0
        [Forms]![Nota de Pedido]![bpos].Caption = "PEDIDOSYA"
        [Forms]![Nota de Pedido]![bpos].Enabled = False
        [Forms]![Nota de Pedido]![Comando68].Visible = False ' Ocultar menu de clientes normal\
        [Forms]![Nota de Pedido]![Comando1204].Visible = True ' mostrar menu de pedidosya normal\
        [Forms]![Nota de Pedido]![Etiqueta1386].Visible = False ' rotulo de eleccion de motorizado
        [Forms]![Nota de Pedido]![CodMotorizado].Visible = False ' lista de eleccion de motorizado
        
        [Forms]![Nota de Pedido]![Comando1669].Visible = False ' Mostrar bton de datos de envio
        [Forms]![Nota de Pedido]![botonaprobar].Visible = False ' boton de aprobar
        
        [Forms]![Nota de Pedido]![logopedidosya1].Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\PedidosYa.ico"
        [Forms]![Nota de Pedido]![logopedidosya2].Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\PedidosYa.ico"
        [Forms]![Nota de Pedido]![logopedidosya1].Visible = True
        [Forms]![Nota de Pedido]![logopedidosya2].Visible = True
        [Forms]![Nota de Pedido]![tipofactura].BackColor = RGB(237, 28, 36)
        [Forms]![Nota de Pedido]![tipopropina].Visible = False
        '[Forms]![Nota de Pedido]![CodCliente].Enabled = False
        DoCmd.OpenForm "Menu PITAYA Delivery"
    Case 4
        [Forms]![Nota de Pedido]![tipofactura] = "PEDIDO CENTRAL ATENCION AL CLIENTE"
        [Forms]![Nota de Pedido]![pacumulado] = 0
        [Forms]![Nota de Pedido]![TipoDelivery] = 0
        [Forms]![Nota de Pedido]![Delivery] = 0
        [Forms]![Nota de Pedido]![POS] = 0
        [Forms]![Nota de Pedido]![Transferencia] = 0
        [Forms]![Nota de Pedido]![bpos].Caption = "EFECTIVO"
        [Forms]![Nota de Pedido]![bpos].Enabled = True
        [Forms]![Nota de Pedido]![Comando68].Visible = True '  menu de clientes normal\
        [Forms]![Nota de Pedido]![Comando1204].Visible = False ' ocultar menu de pedidoya
        '[Forms]![Nota de Pedido]![Etiqueta1386].Visible = True ' rotulo de eleccion de motorizado
        '[Forms]![Nota de Pedido]![CodMotorizado].Visible = True ' lista de eleccion de motorizado
        
        [Forms]![Nota de Pedido]![Comando1669].Visible = True ' Mostrar bton de datos de envio
        [Forms]![Nota de Pedido]![botonaprobar].Visible = True ' boton de aprobar
        
        [Forms]![Nota de Pedido]![CodMotorizado].Visible = False ' ocultar LOSTA DE MOTORIZADOS
        [Forms]![Nota de Pedido]![Etiqueta1386].Visible = False ' ocultar LOSTA DE MOTORIZADOS
        
        DoCmd.SetWarnings False
        DoCmd.OpenForm "Menu PITAYA Global"
End Select


CrearNotaDePedido = "Creado"
End Function


Function ticketpromediodiario(fech As Date) As Double
On Error GoTo Nulo

Dim canti As Integer
Dim tota As Double
canti = DCount("[TotalGuardado]", "[NotaDePedido]", "[TotalGuardado]<>0 AND [Fecha]=#" & fech & "#")
tota = DSum("[TotalGuardado]", "[NotaDePedido]", "[Fecha]=#" & fech & "#")

ticketpromediodiario = tota / canti
Exit Function

Nulo:
ticketpromediodiario = 0
End Function


Function pedidocontienegrupo(tip As String, pedi As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, Grupos.Tipo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY SubPedido.CodPedido, Grupos.Tipo" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((Grupos.Tipo)='" & tip & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienegrupo = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienegrupo = 0
End Function
Function pedidocontienesubgrupo(codgru As Integer, pedi As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, DBBatidos.CodGrupo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, DBBatidos.CodGrupo" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((DBBatidos.CodGrupo)=" & codgru & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienesubgrupo = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienesubgrupo = 0
End Function

Function pedidocontieneproducto(produ As String, pedi As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, DBBatidos.Nombre, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, DBBatidos.Nombre" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((DBBatidos.Nombre)='" & produ & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontieneproducto = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontieneproducto = 0
End Function
Function pedidocontienepromocion(prom As Integer, pedi As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, SubPedido.CodPromocion, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido" & _
" GROUP BY SubPedido.CodPedido, SubPedido.CodPromocion" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((SubPedido.CodPromocion)=" & prom & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienepromocion = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienepromocion = 0
End Function

Function pedidocontienepromocionxgrupo(prom As Integer, pedi As Long, grupix As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, SubPedido.CodPromocion, DBBatidos.CodGrupo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, SubPedido.CodPromocion, DBBatidos.CodGrupo" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((SubPedido.CodPromocion)=" & prom & ")" & _
" AND ((DBBatidos.CodGrupo)=" & grupix & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienepromocionxgrupo = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienepromocionxgrupo = 0
End Function

Function pedidocontienepromocionxsubgrupo(prom As Integer, pedi As Long, grupix As Integer, subgrpix As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, SubPedido.CodPromocion, DBBatidos.CodGrupo, DBBatidos.CodSubGrupo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, SubPedido.CodPromocion, DBBatidos.CodGrupo, DBBatidos.CodSubGrupo" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((SubPedido.CodPromocion)=" & prom & ")" & _
" AND ((DBBatidos.CodGrupo)=" & grupix & ") AND ((DBBatidos.CodSubGrupo)=" & subgrpix & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienepromocionxsubgrupo = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienepromocionxsubgrupo = 0
End Function

Function pedidocontienepromocionxtipo(prom As Integer, pedi As Long, tipox As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, SubPedido.CodPromocion, Grupos.Tipo, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY SubPedido.CodPedido, SubPedido.CodPromocion, Grupos.Tipo" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((SubPedido.CodPromocion)=" & prom & ")" & _
" AND ((Grupos.Tipo)='" & tipox & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienepromocionxtipo = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienepromocionxtipo = 0
End Function

Function pedidocontienegrupotamano(grup As Integer, tama As String, pedi As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodPedido, DBBatidos.CodGrupo, DBBatidos.Medida, Sum(SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY SubPedido.CodPedido, DBBatidos.CodGrupo, DBBatidos.Medida" & _
" HAVING (((SubPedido.CodPedido)=" & pedi & ") AND ((DBBatidos.CodGrupo)=" & grup & ") AND ((DBBatidos.Medida)='" & tama & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienegrupotamano = rst("SumaDeCantidad")
rst.Close
Exit Function

Nulo:
pedidocontienegrupotamano = 0
End Function

Function codigoventatipoenvaseproducto(bati As String) As String
Dim grup As Integer
Dim medi As String
grup = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & bati & "'")
medi = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & bati & "'")

Select Case grup
    Case 1, 2, 3, 4, 8, 16, 24
        Select Case medi
            Case "Mediano"
                codigoventatipoenvaseproducto = "ENV001"
            Case "Gigantona"
                codigoventatipoenvaseproducto = "ENV002"
            Case "Kid"
                codigoventatipoenvaseproducto = "ENV007"
            Case Else
                MsgBox "El Batido no tiene definido su tamano, confirmar con area tecnica"
                codigoventatipoenvaseproducto = ""
        End Select
    Case 6
        codigoventatipoenvaseproducto = "ENV003"
    Case 14
        codigoventatipoenvaseproducto = "ENV006"
    Case 17
        codigoventatipoenvaseproducto = "ENV004"
    Case 18, 19
        codigoventatipoenvaseproducto = "ENV005"
    Case Else
        MsgBox "No se ha encontrado ningun envase de llevar para este producto"
        codigoventatipoenvaseproducto = ""
End Select

End Function


Function statusnotaddepedido(pedi As Long) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.CodPedido, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" FROM Delivery INNER JOIN NotaDePedido ON Delivery.CodDelivery = NotaDePedido.Delivery" & _
" WHERE (((NotaDePedido.CodPedido)=" & pedi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Select Case True

    Case rst("POS") = 0 And rst("Comision") = 0 And rst("Transferencia") = 0
        statusnotaddepedido = "EFECTIVO"
    Case rst("POS") <> 0 And rst("Comision") <> 0 And rst("Transferencia") = 0
        statusnotaddepedido = "PEDIDOSYA"
    Case rst("POS") <> 0 And rst("Comision") = 0 And rst("Transferencia") = 0
        statusnotaddepedido = "POS"
    Case rst("POS") <> 0 And rst("Comision") = 0 And rst("Transferencia") <> 0
        statusnotaddepedido = "TRANSFERENCIA"
    Case Else
        statusnotaddepedido = "OTROS"
End Select


rst.Close
Exit Function

Nulo:
statusnotaddepedido = "OTROS"
End Function

Function statusnotaddepedidodirecto(dPOS As Integer, dDelivery As Integer, dTransferencia As Integer) As String
On Error GoTo Nulo
Dim dComision As Integer

dComision = DLookup("[Comision]", "[Delivery]", "[CodDelivery]=" & dDelivery)

Select Case True

    Case dPOS = 0 And dComision = 0 And dTransferencia = 0
        statusnotaddepedidodirecto = "EFECTIVO"
    Case dPOS <> 0 And dComision <> 0 And dTransferencia = 0
        statusnotaddepedidodirecto = "PEDIDOSYA"
    Case dPOS <> 0 And dComision = 0 And dTransferencia = 0
        statusnotaddepedidodirecto = "POS"
    Case dPOS <> 0 And dComision = 0 And dTransferencia <> 0
        statusnotaddepedidodirecto = "TRANSFERENCIA"
    Case Else
        statusnotaddepedidodirecto = "OTROS"
End Select


Exit Function

Nulo:
statusnotaddepedidodirecto = "OTROS"
End Function


Function statusnotaddepedidodeliverycentral(pedi As Long) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedidoDeliveryCentral.CodPedido, NotaDePedidoDeliveryCentral.POS, Delivery.Comision, NotaDePedidoDeliveryCentral.Transferencia" & _
" FROM Delivery INNER JOIN NotaDePedidoDeliveryCentral ON Delivery.CodDelivery = NotaDePedidoDeliveryCentral.Delivery" & _
" WHERE (((NotaDePedidoDeliveryCentral.CodPedido)=" & pedi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Select Case True

    Case rst("POS") = 0 And rst("Comision") = 0 And rst("Transferencia") = 0
        statusnotaddepedidodeliverycentral = "EFECTIVO"
    Case rst("POS") <> 0 And rst("Comision") <> 0 And rst("Transferencia") = 0
        statusnotaddepedidodeliverycentral = "PEDIDOSYA"
    Case rst("POS") <> 0 And rst("Comision") = 0 And rst("Transferencia") = 0
        statusnotaddepedidodeliverycentral = "POS"
    Case rst("POS") <> 0 And rst("Comision") = 0 And rst("Transferencia") <> 0
        statusnotaddepedidodeliverycentral = "TRANSFERENCIA"
    Case Else
        statusnotaddepedidodeliverycentral = "OTROS"
End Select


rst.Close
Exit Function

Nulo:
statusnotaddepedidodeliverycentral = "OTROS"
End Function

Function montosotrosmediosdiaguardado(fech As Date) As Long

'Monto de total facturado por otros medios solamente de un dia especifico de pedidos no anulados

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'Busqueda de monto de pos del dia
miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS," & _
" Delivery.Comision, NotaDePedido.Transferencia, Sum(NotaDePedido.TotalGuardado) AS Monto" & _
" FROM NotaDePedido INNER JOIN Delivery ON NotaDePedido.Delivery = Delivery.CodDelivery" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.POS, Delivery.Comision, NotaDePedido.Transferencia" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & fech & "#)" & _
" AND ((NotaDePedido.POS)=0) AND ((Delivery.Comision)<>0) AND ((NotaDePedido.Transferencia)=0))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montosotrosmediosdiaguardado = rst("Monto")
rst.Close

Exit Function

NoPos:
montosotrosmediosdiaguardado =  0

End Function

Function pedidocontienebatidoolimonada(pedi As Long) As Integer

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'si el peiddo tien producto de batido o liminoada
miSQL = "SELECT SubPedido.CodPedido, [DBBatidos]![CodGrupo]=1 Or [DBBatidos]![CodGrupo]=2 Or [DBBatidos]![CodGrupo]=3 Or [DBBatidos]![CodGrupo]=8 Or [DBBatidos]![CodGrupo]=16 Or [DBBatidos]![CodGrupo]=24 AS grup" & _
" FROM DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido" & _
" WHERE (((SubPedido.CodPedido) = " & pedi & "))" & _
" ORDER BY [DBBatidos]![CodGrupo]=1 Or [DBBatidos]![CodGrupo]=2 Or [DBBatidos]![CodGrupo]=3 Or [DBBatidos]![CodGrupo]=8 Or [DBBatidos]![CodGrupo]=16 Or [DBBatidos]![CodGrupo]=24"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienebatidoolimonada = rst("grup")
rst.Close

Exit Function

NoPos:
pedidocontienebatidoolimonada =  0

End Function

Function pedidocontienebowlowaffle(pedi As Long) As Integer

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'si el peiddo tien producto de batido o liminoada
miSQL = "SELECT SubPedido.CodPedido, [DBBatidos]![CodGrupo]=6 Or [DBBatidos]![CodGrupo]=14 AS grup" & _
" FROM DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido" & _
" WHERE (((SubPedido.CodPedido) = " & pedi & "))" & _
" ORDER BY [DBBatidos]![CodGrupo]=6 Or [DBBatidos]![CodGrupo]=14"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienebowlowaffle = rst("grup")
rst.Close

Exit Function

NoPos:
pedidocontienebowlowaffle =  0

End Function
Function pedidocontienecuponusado(pedi As Long) As Integer

On Error GoTo NoPos
Dim rst As DAO.Recordset
Dim miSQL As String

'si el peiddo tien producto de batido o liminoada
miSQL = "SELECT SubPedido.CodPedido, [SubPedido]![CodPromocion]=123 Or [SubPedido]![CodPromocion]=124 Or [SubPedido]![CodPromocion]=125 AS prom" & _
" FROM DBBatidos INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido" & _
" WHERE (((SubPedido.CodPedido) = " & pedi & "))" & _
" ORDER BY [SubPedido]![CodPromocion]=123 Or [SubPedido]![CodPromocion]=124 Or [SubPedido]![CodPromocion]=125"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
pedidocontienecuponusado = rst("prom")
rst.Close

Exit Function

NoPos:
pedidocontienecuponusado =  0

End Function



Function porcioneanexadaasubreceta(codsubrec As Long) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodSubReceta, SubReceta.codporcion FROM SubReceta WHERE (((SubReceta.CodSubReceta)=" & codsubrec & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
porcioneanexadaasubreceta = rst("codporcion")

rst.Close

Exit Function

Nulo:
porcioneanexadaasubreceta = 0
End Function
Function nombreinsumoreceta(codsubrec As Long) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodSubReceta, SubReceta.codporcion, DBIngredientes.Nombre," & _
" DBIngredientes.NombreSinProcesar, DBIngredientes.NombreProcesado, DBIngredientes.NombreReceta," & _
" DBIngredientes.presentacionpreparacion, DBIngredientes.Unidad, DBIngredientes.conversionpreparacion" & _
" FROM SubReceta INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((SubReceta.CodSubReceta)=" & codsubrec & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If porcioneanexadaasubreceta(codsubrec) = 0 Then ' no tinee procion anexada
    If rst("conversionpreparacion") <> 1 Then ' la presentacion de prearacion es distinta a 1 , tiene rpesetacion de preparacion personalizada
        If rst("Nombre") = "Fresa Fresca" Or rst("Nombre") = "Naranja Dulce" Then        '||ANEXAREMOS LA EXCEPCION PARA LA conversiones que no son porciones pero se porcionan en tineda y tiene presentacion distitna
            nombreinsumoreceta = Left(rst("NombreReceta"), 10) & " (porcion " & LCase(rst("presentacionpreparacion")) & ")"
        Else
            nombreinsumoreceta = rst("NombreReceta") & " " & LCase(rst("presentacionpreparacion")) '
        End If
    Else
        nombreinsumoreceta = rst("NombreReceta") & " (" & LCase(rst("Unidad")) & ")"
    End If
    
Else ' es porcion y tienen porcion anexado  ||ANEXAREMOS LA EXCEPCION PARA LA PITAYA SE MUESTRE PORCIONES||
    If rst("Nombre") = "Pitaya" Or rst("Nombre") = "Melon" Or rst("Nombre") = "Fresa Congelada" Or rst("Nombre") = "Mora y Arandano" Or rst("Nombre") = "Mango" Or rst("Nombre") = "Papaya" Or rst("Nombre") = "Piña" Or rst("Nombre") = "Sandia" Or rst("Nombre") = "Jugo de Limon" Then        '||ANEXAREMOS LA EXCEPCION PARA LA PITAYA SE MUESTRE PORCIONES||
        nombreinsumoreceta = Left(rst("NombreProcesado") & " " & DLookup("[Linea]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion")), 10) & " (porcion " & DLookup("[Capacidad]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion")) & ")" 'nombreproductoporcionreceta(rst("codporcion"))
    Else
        nombreinsumoreceta = rst("NombreProcesado") & " " & DLookup("[Linea]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion")) & " (unid)" 'nombreproductoporcionreceta(rst("codporcion"))
    
    End If
End If

rst.Close

Exit Function

Nulo:
nombreinsumoreceta = ""
End Function

Function nombreinsumorecetasinunidadmedida(codsubrec As Long) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodSubReceta, SubReceta.codporcion, DBIngredientes.Nombre , DBIngredientes.NombreSinProcesar, DBIngredientes.NombreProcesado, DBIngredientes.NombreReceta," & _
" DBIngredientes.presentacionpreparacion, DBIngredientes.Unidad, DBIngredientes.conversionpreparacion" & _
" FROM SubReceta INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((SubReceta.CodSubReceta)=" & codsubrec & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If porcioneanexadaasubreceta(codsubrec) = 0 Then ' no tinee procion anexada
    If rst("conversionpreparacion") <> 1 Then ' la presentacion de prearacion es distinta a 1 , tiene rpesetacion de preparacion personalizada
        nombreinsumorecetasinunidadmedida = rst("NombreReceta")
    Else
        nombreinsumorecetasinunidadmedida = rst("NombreReceta")
    End If
    
Else
    nombreinsumorecetasinunidadmedida = rst("NombreProcesado") & " " & DLookup("[Linea]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion"))  'nombreproductoporcionreceta(rst("codporcion"))
End If

rst.Close

Exit Function

Nulo:
nombreinsumorecetasinunidadmedida = ""
End Function

Function unidadmedidainsumoreceta(codsubrec As Long) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodSubReceta, SubReceta.codporcion, DBIngredientes.Nombre , DBIngredientes.NombreSinProcesar, DBIngredientes.NombreProcesado, DBIngredientes.NombreReceta," & _
" DBIngredientes.presentacionpreparacion, DBIngredientes.Unidad, DBIngredientes.conversionpreparacion" & _
" FROM SubReceta INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((SubReceta.CodSubReceta)=" & codsubrec & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If porcioneanexadaasubreceta(codsubrec) = 0 Then ' no tinee procion anexada
    If rst("conversionpreparacion") <> 1 Then ' la presentacion de prearacion es distinta a 1 , tiene rpesetacion de preparacion personalizada
        unidadmedidainsumoreceta = LCase(rst("presentacionpreparacion"))
    Else
        unidadmedidainsumoreceta = LCase(rst("Unidad"))
    End If
    
Else
    unidadmedidainsumoreceta = "unid"   'nombreproductoporcionreceta(rst("codporcion"))
End If

rst.Close

Exit Function

Nulo:
unidadmedidainsumoreceta = ""
End Function

Function cantidadinsumoreceta(codsubrec As Long, cantidadrelativa As Double) As String
On Error GoTo Nulo

cantidadinsumoreceta = fraccion(cantidadinsumorecetaantesconvertirfraccion(codsubrec) * cantidadrelativa)

Exit Function

Nulo:
cantidadinsumoreceta = ""
End Function

Function cantidadinsumorecetaantesconvertirfraccion(codsubrec As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubReceta.CodSubReceta, SubReceta.codporcion, SubReceta.Cantidad," & _
" DBIngredientes.conversionpreparacion" & _
" FROM SubReceta INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((SubReceta.CodSubReceta)=" & codsubrec & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If porcioneanexadaasubreceta(codsubrec) = 0 Then 'no tiene procion anexada
    If rst("conversionpreparacion") <> 1 Then ' la presentacion de prearacion es distinta a 1 , tiene rpesetacion de preparacion personalizada
        cantidadinsumorecetaantesconvertirfraccion = rst("Cantidad") / rst("conversionpreparacion")
    Else
        cantidadinsumorecetaantesconvertirfraccion = rst("Cantidad")
    End If
    
Else
    cantidadinsumorecetaantesconvertirfraccion = rst("Cantidad") / DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion"))
End If

rst.Close

Exit Function

Nulo:
cantidadinsumorecetaantesconvertirfraccion = 0
End Function

Function comandalistadatos(bati As String, Data As Integer, CodSubPedido As Long)
'data:
'1: tipos
'2: Nombre
'3: Cantidad
'4: base o no letra po letra

'On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim agregado As String
Dim tiposalto As String
Dim tipogruposalto As String
Dim saltodelinea As String
Dim cuenta As Integer
Dim arrayendulzante As Variant
arrayendulzante = arrayendulzanteagregadorecetacantidad(CodSubPedido) 'codigo + cantidad

'miSQL = "SELECT TipoIngredientesReceta.Orden, SubReceta.ordenreceta, TipoIngredientesReceta.GrupoTipoReceta, SubReceta.Tipo, DBIngredientes.Tipo," & _
" nombreinsumoreceta([SubReceta]![CodSubReceta]) AS nombrecondi, cantidadinsumoreceta([SubReceta]![CodSubReceta]) AS cantidadcondi," & _
" SubReceta.CodBatido," & _
" IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))) AS empaquealfinal" & _
" FROM TipoIngredientesReceta INNER JOIN (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" ON TipoIngredientesReceta.CodTipoIngredientesReceta = SubReceta.Tipo" & _
" WHERE (((SubReceta.CodBatido)='" & bati & "')" & _
" AND ((IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))))=0))" & _
" ORDER BY TipoIngredientesReceta.Orden, SubReceta.ordenreceta"

'miSQL = "SELECT TipoIngredientesReceta.Orden, SubReceta.ordenreceta, TipoIngredientesReceta.GrupoTipoReceta, SubReceta.Tipo, DBIngredientes.Tipo as TipoIngre," & _
" nombreinsumoreceta([SubReceta]![CodSubReceta]) AS nombrecondi," & _
" cantidadinsumoreceta([SubReceta]![CodSubReceta]) AS cantidadcondi," & _
" SubReceta.CodBatido," & _
" IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))) AS empaquealfinal" & _
" FROM TipoIngredientesReceta INNER JOIN (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente)" & _
" ON TipoIngredientesReceta.CodTipoIngredientesReceta = SubReceta.Tipo" & _
" WHERE (((SubReceta.CodBatido)='" & bati & "')" & _
" AND ((IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))))=0))" & _
" UNION ALL " & _
" SELECT 13 AS Orden, 1 AS ordenreceta, 'otro' AS GrupoTipoReceta, 'E' AS Tipo, 'E' AS TipoIngre," & _
" endulzanteagregadorecetanombre(" & CodSubPedido & ") AS nombrecondi, fraccion(endulzanteagregadorecetacantidad(" & CodSubPedido & ")) AS cantidadcondi, " & _
" '" & bati & "' AS CodBatido," & _
" 1 AS empaquealfinal" & _
" FROM (SELECT TOP 1 1 FROM MSysObjects) AS Dummy" & _
" ORDER BY Orden, ordenreceta"
'(SELECT TOP 1 1 FROM MSysObjects) AS Dummy artilugio para tener minimo 1 registro en endulzante
        
miSQL = "SELECT TipoIngredientesReceta.Orden, SubReceta.ordenreceta, " & _
        "TipoIngredientesReceta.GrupoTipoReceta, SubReceta.Tipo, " & _
        "DBIngredientes.Tipo as TipoIngre, " & _
        "nombreinsumoreceta([SubReceta]![CodSubReceta]) AS nombrecondi, " & _
        "cantidadinsumoreceta([SubReceta]![CodSubReceta], " & _
        "IIf(SubReceta.CodBatido = '" & arrayendulzante(0) & "', " & arrayendulzante(1) & ", 1)) AS cantidadcondi, " & _
        "SubReceta.CodBatido, " & _
        "IIf([SubReceta]![Tipo]='P', 1, " & _
        "IIf(IsNull([SubReceta]![codporcion]), 0, " & _
        "InsumoMezclaPorcion([SubReceta]![codporcion]))) AS empaquealfinal " & _
        "FROM (TipoIngredientesReceta " & _
        "INNER JOIN (DBIngredientes " & _
        "INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente) " & _
        "ON TipoIngredientesReceta.CodTipoIngredientesReceta = SubReceta.Tipo) " & _
        "WHERE ((SubReceta.CodBatido='" & bati & "') OR (SubReceta.CodBatido='" & arrayendulzante(0) & "')) " & _
        "AND ((IIf([SubReceta]![Tipo]='P',1,IIf(IsNull([SubReceta]![codporcion]),0,InsumoMezclaPorcion([SubReceta]![codporcion]))))=0)" & _
        "ORDER BY Orden, ordenreceta"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
comandalistadatos = ""
saltodelinea = ""
tiposalto = rst("Tipo")
tipogruposalto = rst("GrupoTipoReceta")
cuenta = 1

For cuenta = 1 To cantli

    
    If tipogruposalto <> rst("GrupoTipoReceta") Then
        Select Case Data
            Case 1
                'saltodelinea = "---"
                saltodelinea = "__"
            Case 2
                'saltodelinea = "---------------------------------------"
                saltodelinea = "_________________________"
            Case 3
                'saltodelinea = "---------"
                saltodelinea = "______"
            Case 4
                'saltodelinea = "---"
                saltodelinea = "__"
        End Select
        comandalistadatos = comandalistadatos & saltodelinea & vbCrLf
        tipogruposalto = rst("GrupoTipoReceta")
        tiposalto = rst("Tipo")

    Else
        If tiposalto <> rst("Tipo") Then
            comandalistadatos = comandalistadatos & vbCrLf
            tiposalto = rst("Tipo")
        End If
    End If
    
    Select Case Data
        Case 1
            agregado = rst("Tipo")
        Case 2
            agregado = Left(rst("nombrecondi"), 25)
        Case 3
            agregado = rst("cantidadcondi")
        Case 4
            If cuenta = 1 And tipogruposalto = "base" Then
                agregado = "B"
            End If
        Case Else
            agregado = ""
    End Select
    comandalistadatos = comandalistadatos & agregado & vbCrLf
    agregado = ""
    rst.MoveNext
Next cuenta

rst.Close

Exit Function

AgainAgain:
comandalistadatos = ""
End Function

Function comandaendulzantesdatos(subpedi As Long, Data As Integer)
'data:
'1: tipos
'2: Nombre
'3: Cantidad

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim agregado As String
Dim bati As String
Dim condi As Double
bati = DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & subpedi)

miSQL = "SELECT DBIngredientes.Tipo, DBIngredientes.CodIngrediente," & _
" [DBIngredientes]![Nombre] & ' (' & [DBIngredientes]![Unidad] & ')' As name" & _
" FROM DBIngredientes WHERE (((DBIngredientes.Tipo)='Endulzantes'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
comandaendulzantesdatos = ""

For I = 1 To cantli
    If DLookup("[Endulzante]", "[Grupos]", "[CodGrupo]=" & DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & bati & "'")) > 0 Then
        'batido con enulzante  agrega los endulzantes aparte
        condi = CantidadIngrediente(bati, rst("CodIngrediente")) + CantidadIngredienteAnexadoPedido(subpedi, rst("CodIngrediente"))
        If condi <> 0 Then
            Select Case Data
                Case 1
                    agregado = "B"
                Case 2
                    agregado = Left(rst("name"), 25)
                Case 3
                    agregado = fraccion(condi)
                Case Else
                    agregado = ""
            End Select
            comandaendulzantesdatos = comandaendulzantesdatos & agregado & vbCrLf
        End If
    Else
        ' producto sin endulzante no suma lineas
    End If
    
    rst.MoveNext
Next I

rst.Close

Exit Function

AgainAgain:
comandaendulzantesdatos = ""
End Function

Function CantidadIngredienteAnexadoPedido(subpedi As Long, ingre As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la cantidad del ingrediente especifico en la recetas que estana nexadas a un pedido
miSQL = "SELECT SubPedido.VInculo, SubReceta.CodIngrediente, Sum(SubReceta.Cantidad*SubPedido.Cantidad) AS SumaDeCantidad" & _
" FROM (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN SubPedido ON DBBatidos.CodBatido = SubPedido.CodBatido" & _
" GROUP BY SubPedido.VInculo, SubReceta.CodIngrediente" & _
" HAVING (((SubPedido.VInculo)=" & subpedi & ") AND ((SubReceta.CodIngrediente)='" & ingre & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadIngredienteAnexadoPedido = rst("SumaDeCantidad")
rst.Close
     
Exit Function

Nulo:
CantidadIngredienteAnexadoPedido = 0

End Function

Function existesolicitudanulacionpedido(codped As Long) As Integer
'1 existe
'0 no existe
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT AnulacionPedidos.CodPedido, 1 AS cont" & _
" FROM AnulacionPedidos WHERE (((AnulacionPedidos.CodPedido)=" & codped & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
existesolicitudanulacionpedido = rst("cont")
rst.Close
     
Exit Function

Nulo:
existesolicitudanulacionpedido = 0

End Function

Function statusanuladopedido(codped As Long) As String
If DLookup("[Anulado]", "[NotaDePedido]", "[CodPedido]=" & codped) <> 0 Then
    statusanuladopedido = "Anulado"
Else
    If existesolicitudanulacionpedido(codped) = 0 Then ' no existe solicitud de anualcion
        statusanuladopedido = "Activo"
    Else
        If DLookup("[Status]", "[AnulacionPedidos]", "[CodPedido]=" & codped) = 0 Then
            statusanuladopedido = "Pendiente"
        Else
            statusanuladopedido = "Anulado"
        End If
    End If
End If
End Function

Function motivosolicitudanulacionpedido(codped As Long) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT AnulacionPedidos.CodPedido, AnulacionPedidos.Motivo FROM AnulacionPedidos" & _
" WHERE (((AnulacionPedidos.CodPedido)=" & codped & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
motivosolicitudanulacionpedido = rst("Motivo")
rst.Close
     
Exit Function

Nulo:
motivosolicitudanulacionpedido = ""

End Function

Function ultimopedidofacturado() As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedido.CodPedido FROM NotaDePedido ORDER BY NotaDePedido.CodPedido DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimopedidofacturado = rst("CodPedido")
rst.Close
     
Exit Function

Nulo:
ultimopedidofacturado = 0

End Function

Function ultimosubpedidofacturado() As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPedido.CodSubPedido FROM SubPedido ORDER BY SubPedido.CodSubPedido DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimosubpedidofacturado = rst("CodSubPedido")
rst.Close
     
Exit Function

Nulo:
ultimosubpedidofacturado = 0

End Function

Function endulzanteagregadorecetanombre(pedidox As Long) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.Nombre, SubPedido.Vinculo, DBIngredientes.Unidad," & _
" [DBIngredientes]![Nombre] & ' (' & [DBIngredientes]![Unidad] & ')' AS Expr1" & _
" FROM ((SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((DBBatidos.Nombre) Like '*Endulzante') AND ((SubPedido.Vinculo)=" & pedidox & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
endulzanteagregadorecetanombre = rst("Expr1")
If rst("Expr1") = "Miel (gr)" Then
    endulzanteagregadorecetanombre = "Miel (oz)"
End If
rst.Close
     
Exit Function

Nulo:
endulzanteagregadorecetanombre = ""

End Function

Function endulzanteagregadorecetacantidad(pedidox As Long) As Double
'pedidox = codsubpedido
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.Nombre, SubPedido.Vinculo, SubPedido.Cantidad" & _
" FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" WHERE (((DBBatidos.Nombre) Like '*Endulzante') AND ((SubPedido.Vinculo)=" & pedidox & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
endulzanteagregadorecetacantidad = rst("Cantidad")
rst.Close
     
Exit Function

Nulo:
endulzanteagregadorecetacantidad = 0

End Function

Function arrayendulzanteagregadorecetacantidad(pedidox As Long) As Variant
    'codigo de batido(endulzante  + cantidad agregada
    On Error GoTo Nulo
    Dim rst As DAO.Recordset
    Dim miSQL As String
    Dim resultado(1) As Variant
    
    miSQL = "SELECT DBBatidos.Nombre, DBBatidos.CodBatido, SubPedido.Vinculo, SubPedido.Cantidad" & _
            " FROM SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
            " WHERE (((DBBatidos.Nombre) Like '*Endulzante') AND ((SubPedido.Vinculo)=" & pedidox & "))"
    
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    
    If Not rst.EOF Then
        rst.MoveFirst
        ' Almacenar CodBatido como string en posición 0
        resultado(0) = CStr(rst("CodBatido"))
        ' Almacenar Cantidad como double en posición 1
        resultado(1) = CDbl(rst("Cantidad"))
    Else
        ' Si no hay registros, devolver array vacío
        resultado(0) = ""
        resultado(1) = 0
    End If
    
    rst.Close
    arrayendulzanteagregadorecetacantidad = resultado
    Exit Function

Nulo:
    ' En caso de error, devolver array con valores por defecto
    resultado(0) = ""
    resultado(1) = 0
    arrayendulzanteagregadorecetacantidad = resultado
End Function

Sub resetearpromocionespedidocompleto(pedid As Long)
DoCmd.SetWarnings False

    'DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.CodPromocion = 5 " & _
                 "WHERE SubPedido.CodPedido = " & pedid & _
                 " AND DLookUp('[usointerno]','[DBPromociones]','[CodPromocion]=' & SubPedido.CodPromocion)=0"
    DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.CodPromocion = 5 " & _
                 "WHERE SubPedido.CodPedido = " & pedid & _
                 " AND (DLookUp('[usointerno]','[DBPromociones]','[CodPromocion]=' & SubPedido.CodPromocion)=0" & _
                 " AND DLookUp('[CodGrupo]','[DBBatidos]','[CodBatido]=''' & SubPedido.CodBatido & '''')<>20)"

DoCmd.SetWarnings True
End Sub

Public Sub EliminarSubPedidoRecursivo(ByVal codRaiz As Long, ByVal pedi As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim hijos() As Long
    Dim n As Integer
    Dim I As Integer

    Set db = CurrentDb

    ' 1) Leer todos los hijos y guardarlos en memoria ANTES de borrar nada
    sql = "SELECT CodSubpedido FROM SubPedido WHERE Vinculo = " & codRaiz & " AND CodPedido = " & pedi
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    n = 0
    If Not rs.EOF Then
        rs.MoveLast   ' fuerza a materializar todo el snapshot
        n = rs.RecordCount
        ReDim hijos(1 To n)
        rs.MoveFirst
        I = 1
        Do While Not rs.EOF
            hijos(I) = rs!CodSubPedido
            I = I + 1
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing

    ' 2) Ahora sí, recursar sobre el array ya cerrado el recordset
    For I = 1 To n
        EliminarSubPedidoRecursivo hijos(I), pedi
    Next I

    ' 3) Borrar el nodo actual (y cualquier vínculo residual)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM SubPedido WHERE Vinculo = " & codRaiz & " AND CodPedido = " & pedi
    DoCmd.RunSQL "DELETE * FROM SubPedido WHERE CodSubpedido = " & codRaiz & " AND CodPedido = " & pedi
    DoCmd.SetWarnings True

    Set db = Nothing
End Sub
