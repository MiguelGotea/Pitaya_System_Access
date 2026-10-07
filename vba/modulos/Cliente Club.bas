' ==========================================================
' Modulo  : Cliente Club
' Tipo    : 1  |  Lineas: 917
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database
Function PuntosGlobales2(codc As Integer) As Double
'cantidad de puntos SUMA en base a una consulta del MIXED donde esta subpedido y nota de pedido de cada local UNIDO

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim rutaDB As String

If codc = 0 Or codc = 1243 Or codc = 2500 Then
    PuntosGlobales2 = 0
    Exit Function
End If

'LLAMADO DE CONSULTA de subpedido y nota de pedido donde se extraera toda la info del codigo
rutaDB = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"

    'LLAMADO DE CONSULT de subepido ynoa de pedido donde se extraera toda la info del codigo
    'miSQL = "SELECT DetallePedidoClub.CodCliente, Sum(PuntosClubProducto([CodPromocion],[DBBatidos].[CodBatido],[Fecha])*[Cantidad]) AS Puntos" & _
    " FROM DetallePedidoClub" & _
    " INNER JOIN DBBatidos ON DetallePedidoClub.CodBatido = DBBatidos.CodBatido" & _
    " IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
    " GROUP BY DetallePedidoClub.CodCliente HAVING (((DetallePedidoClub.CodCliente)=" & codc & "))"
miSQL = "SELECT dpc.CodCliente, Sum(PuntosClubProducto(dpc.CodPromocion, b.CodBatido, dpc.Fecha) * dpc.Cantidad) AS Puntos " & _
        "FROM [;DATABASE=" & rutaDB & "].DetallePedidoClub AS dpc, " & _
        "DBBatidos AS b " & _
        "WHERE dpc.CodBatido = b.CodBatido " & _
        "AND dpc.CodCliente = " & codc & " " & _
        "GROUP BY dpc.CodCliente"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Not rst.EOF Then
    PuntosGlobales2 = rst("Puntos") + InicialesCliente(codc)
Else
    PuntosGlobales2 = InicialesCliente(codc)
End If

rst.Close
Set rst = Nothing
Exit Function

Nulo:
PuntosGlobales2 = InicialesCliente(codc)
End Function

Function PuntosGlobales3(codc As Integer) As Double
'cantidad de puntos SUMA en base a la descarga desde el host y la tabla local

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim rutaDB As String

If codc = 0 Or codc = 1243 Or codc = 2500 Then
    PuntosGlobales3 = 0
    Exit Function
End If

miSQL = "SELECT CodCliente, Anulado, Sum(Puntos*Cantidad) AS SumaDePuntos" & _
" FROM VentasGlobalesAccessCSVFiltradoClienteInternoExterno" & _
" GROUP BY CodCliente, Anulado" & _
" HAVING (((CodCliente)=" & codc & ") AND ((Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Not rst.EOF Then
    PuntosGlobales3 = rst("SumaDePuntos") + InicialesCliente2(codc)
Else
    PuntosGlobales3 = InicialesCliente2(codc)
End If

rst.Close
Set rst = Nothing
Exit Function

Nulo:
PuntosGlobales3 = InicialesCliente2(codc)
End Function

Function PuntosGlobales4(codc As Long) As Double
On Error GoTo Nulo

If codc = 0 Or codc = 1243 Or codc = 2500 Then
    PuntosGlobales4 = 0
    Exit Function
End If

Dim Cliente As Variant
Cliente = DatosClubHost(codc, codigoLocal())

' puntos en el host + puntos locales  + puntos inciiales en el hsot + puntos inciales lcoales
PuntosGlobales4 = Cliente(3) + PuntosLocalesClub(codc) + Cliente(4) + InicialesClienteLocal(codc)
Exit Function

Nulo:
PuntosGlobales4 = 0
End Function
Function DatosClienteClubGlobal(codc As Long) As Variant
On Error GoTo Nulo

Dim resultado(0 To 4) As Variant
Dim Cliente As Variant
Dim nombreCliente As String
Dim puntosinicialestotal As Double

' Inicializar valores
resultado(0) = 0 ' Double para los puntos
resultado(1) = "" ' String para el nombre
resultado(2) = 1 ' entero 0 si no existe, 1 si existe
resultado(3) = 0 ' cantidad de puntos iniciales

If codc = 0 Or codc = 1243 Or codc = 2500 Then
    DatosClienteClubGlobal = resultado
    Exit Function
End If

Cliente = DatosClubHost(codc, codigoLocal()) 'Datos de clinte exceptuando los locales
puntosinicialestotal = IIf(Cliente(4) = 0, InicialesClienteLocal(codc), Cliente(4)) ' si no encuentra puntos afuera entonces busca local
resultado(3) = puntosinicialestotal
resultado(0) = Cliente(3) + PuntosLocalesClub(codc) + puntosinicialestotal

' Determinar el nombre del cliente según las condiciones
If Cliente(0) = 0 Then 'no existe cleinte en el host
    resultado(1) = Nz(DLookup("[Nombre]", "[ClientesClub]", "[CodCliente]=" & codc), "") ' busca local
Else
    ' Usar cliente(2) como nombre
    resultado(1) = Cliente(2)
End If

If Cliente(0) = 0 And Nz(DLookup("[CodCliente]", "[ClientesClub]", "[CodCliente]=" & codc), 0) = 0 Then
    resultado(2) = 0
End If

DatosClienteClubGlobal = resultado
Exit Function

Nulo:
' En caso de error, devolver array con valores por defecto
resultado(0) = 0
resultado(1) = ""
resultado(2) = 0
resultado(3) = 0
DatosClienteClubGlobal = resultado
End Function
 
Function PuntosLocalesClub(codc As Long) As Double
'puntos acu,ulados en local base de datos de sistema
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim rutaDB As String

miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.CodCliente, Sum(SubPedido.Puntos*SubPedido.Cantidad) AS SumaDePuntos" & _
" FROM SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.CodCliente" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.CodCliente)=" & codc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PuntosLocalesClub = rst("SumaDePuntos")

rst.Close
Set rst = Nothing
Exit Function

Nulo:
PuntosLocalesClub = 0
End Function

Function PuntosNotaPedido(codp As Long) As Double
'Cantidad de puntos de una nota de pedido especifico, considerando codpromocion y productos que suman puntos

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de productos dentro de los parametros de codpromocion y sumpuntos para un codpedido especifico
miSQL = "SELECT SubPedido.CodPedido," & _
" Sum([SubPedido]![Cantidad]*PuntosClubProducto([SubPedido]![CodPromocion],[SubPedido]![CodBatido],[NotadePedido]![Fecha])) AS puntos" & _
" FROM SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido" & _
" GROUP BY SubPedido.CodPedido" & _
" HAVING (((SubPedido.CodPedido)=" & codp & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PuntosNotaPedido = rst("puntos")
rst.Close


Exit Function

Nulo:
PuntosNotaPedido = 0

End Function

Sub ActualizarPuntosSubPedidoDePedido(codp As Long)
    ' actualizar puntos de subpedido de un pedido
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    Dim db As DAO.Database
    
    Set db = CurrentDb
    
    strSQL = "UPDATE ((SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido)" & _
            " INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
            " SET SubPedido.Puntos = PuntosClubProducto([SubPedido]![CodPromocion],[SubPedido]![CodBatido],[NotaDePedido]![Fecha])" & _
            " WHERE (((SubPedido.CodPedido)=" & codp & ") AND ((Grupos.SumaPuntos)<>0));"
    
    'Ejecutar la consulta
    db.Execute strSQL, dbFailOnError
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al actualizar puntos del pedido " & codp & ": " & Err.Description, vbCritical

End Sub

Function PuntosClubProducto(codprom As Integer, codbat As String, fech As Date) As Double
'puntos por cada producto acorde a su condicion de tipo producto medida y promocion q tiene
'CodBatido = DBBatidos!codbatido
'Tipo = Grupo!Tipo
'Fecha = NotadePedido!Fecha

'On Error GoTo Nulo

Dim Medida As String
Dim grup As Integer
Dim nombrebat As String
Dim marc As String
Dim subgrup As Integer
Medida = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & codbat & "'")
grup = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & codbat & "'")
subgrup = DLookup("[CodSubGrupo]", "[DBBatidos]", "[CodBatido]='" & codbat & "'")
marc = DLookup("[Marca]", "[DBBatidos]", "[CodBatido]='" & codbat & "'")
nombrebat = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & codbat & "'")

Select Case codprom
    Case 5, 89, 91 ' sin promocion, suma puntos rentencio de 2% y 3% tmb suma
        
        Select Case fech
            Case Is > #7/3/2025#
                Select Case grup
                    Case 1, 2, 3, 4, 8, 16, 24 ' batidos
                        PuntosClubProducto = 1
                    Case 14, 6 ' waffles y bowl
                        PuntosClubProducto = 1.5
                    Case 7
                        If subgrup = 1 Then
                            PuntosClubProducto = 0.8
                        ElseIf subgrup = 2 Then
                            PuntosClubProducto = 1
                        Else
                            PuntosClubProducto = 0
                        End If
                    Case Else
                        PuntosClubProducto = 0
                End Select
                
            Case #1/1/2025# To #7/3/2025#  'se agrego limonadas
                Select Case grup
                    Case 1, 2, 3, 4, 8, 16, 24 ' batidos
                        PuntosClubProducto = 1
                    Case 14, 6 ' waffles y bowl
                        PuntosClubProducto = 1.5
                    Case 7
                        If subgrup = 1 Then
                            PuntosClubProducto = 0.8
                        Else
                            PuntosClubProducto = 0
                        End If
                    Case Else
                        PuntosClubProducto = 0
                End Select
                
            Case #10/18/2024# To #12/31/2024#
                Select Case grup
                    Case 1, 2, 3, 4, 8, 24 ' batidos
                        PuntosClubProducto = 1
                    Case 14, 6 ' waffles y bowl
                        PuntosClubProducto = 1.5
                    Case 7
                        If subgrup = 1 Then
                            PuntosClubProducto = 0.8
                        Else
                            PuntosClubProducto = 0
                        End If
                    Case Else
                        PuntosClubProducto = 0
                End Select
            
            Case #12/21/2022# To #10/17/2024#
                Select Case grup
                    Case 1, 2, 3, 4, 8, 24 ' batidos
                        PuntosClubProducto = 1
                    Case 14, 6 ' waffles y bowl
                        PuntosClubProducto = 1.5
                    Case Else
                        PuntosClubProducto = 0
                End Select
                
            Case #10/16/2017# To #12/20/2022#
               Select Case Medida
                    Case "Kid"
                        PuntosClubProducto = 1
                    Case "Mediano"
                        PuntosClubProducto = 1
                    Case "Gigantona"
                        PuntosClubProducto = 1
                    Case "Bowl"
                        PuntosClubProducto = 1.5
                    Case Else
                        PuntosClubProducto = 0
                End Select
            Case Is < #10/16/2017#
                Select Case Medida
                    Case "Kid", "Mediano", "Gigantona", "Bowl"
                        PuntosClubProducto = 1
                    Case Else
                        PuntosClubProducto = 0
                End Select
        End Select
                
    Case 22 ' canjeado x puntos
    
        Select Case fech
        
            Case Is > #10/5/2026# ' Politica Oct 2026: por tamano, tipo y nombre de producto. Restricciones adicionales en ventana de facturacion

                Select Case grup
                    Case 1, 3, 4, 8 ' Especiales, Clasicos, Estrenos, Saludables
                        Select Case Medida
                            Case "Mediano"
                                PuntosClubProducto = -10
                            Case "Gigantona"
                                If grup = 3 Then
                                    PuntosClubProducto = -11 ' Clasicos 20oz
                                Else
                                    PuntosClubProducto = -13 ' Especiales, Estrenos, Saludables 20oz
                                End If
                            Case Else
                                PuntosClubProducto = 0
                        End Select
                    Case 2, 24 ' Premium y ConProteina
                        Select Case Medida
                            Case "Mediano"
                                PuntosClubProducto = -12
                            Case "Gigantona"
                                PuntosClubProducto = -15
                            Case Else
                                PuntosClubProducto = 0
                        End Select
                    Case 16 ' Limonadas - nuevo en canjeo desde Oct 2026
                        Select Case Medida
                            Case "Mediano"
                                If nombrebat Like "*Fresa*" Then
                                    PuntosClubProducto = -11
                                Else
                                    PuntosClubProducto = -10
                                End If
                            Case "Gigantona"
                                PuntosClubProducto = -13
                            Case Else
                                PuntosClubProducto = 0
                        End Select
                    Case 14 ' Waffles diferenciados por nombre
                        If nombrebat Like "*Clasico*" Then
                            PuntosClubProducto = -15
                        ElseIf nombrebat Like "*Especial*" Then
                            PuntosClubProducto = -16
                        ElseIf nombrebat Like "*Proteina*" Then
                            PuntosClubProducto = -18
                        Else
                            PuntosClubProducto = 0
                        End If
                    Case 6 ' Energy Bowl - Acai ahora SI aplica desde Oct 2026
                        If nombrebat Like "*Acai*" Then
                            PuntosClubProducto = -22
                        ElseIf nombrebat Like "*Dragon*" Then
                            PuntosClubProducto = -20
                        Else ' Fachento, Ometepe y demas
                            PuntosClubProducto = -18
                        End If
                    Case 7 ' Pitaya Store por nombre de producto
                        Select Case subgrup
                            Case 1 ' Semillas pequenas y mixes
                                If nombrebat Like "*Almendras*" Or nombrebat Like "*Pistachos*" Then
                                    PuntosClubProducto = -8
                                Else ' Maranon, Mix Cacao Energy, Mix Supremo, Mix Tropical
                                    PuntosClubProducto = -12
                                End If
                            Case 2 ' Frascos grandes
                                If nombrebat Like "*Semilla de Cacao*" Then
                                    PuntosClubProducto = -29
                                Else ' Granola Grande, Cocoa en Polvo Grande
                                    PuntosClubProducto = -21
                                End If
                            Case 3 ' Galletas de Avena
                                PuntosClubProducto = -2.5
                            Case Else
                                PuntosClubProducto = 0
                        End Select
                    Case Else
                        PuntosClubProducto = 0
                End Select
                'MsgBox "Puntos canjeados: " & PuntosClubProducto

            Case #7/4/2025# To #10/5/2026# ' Politica jul 2025 - oct 2026. Las restricciones se hacen en la ventana de facturacion

                Select Case grup
                    Case 2
                        PuntosClubProducto = -12
                    Case 1, 3, 4, 8
                        PuntosClubProducto = -10
                    Case 24
                        PuntosClubProducto = -12
                    Case 14
                        PuntosClubProducto = -15
                    Case 6
                        PuntosClubProducto = -18
                    Case 7
                        Select Case subgrup
                            Case 1
                                PuntosClubProducto = -8
                            Case 3
                                PuntosClubProducto = -2.5
                            Case Else
                                PuntosClubProducto = 0
                        End Select
                    Case Else
                        PuntosClubProducto = 0
                End Select

                
            Case #10/18/2024# To #7/3/2025# ' Politica oct 2024 - jul 2025. Las restricciones se hacen en la ventana de facturacion

                Select Case grup
                    Case 2
                        PuntosClubProducto = -12
                    Case 1, 3, 4, 8
                        PuntosClubProducto = -10
                    Case 24
                        PuntosClubProducto = -12
                    Case 14
                        PuntosClubProducto = -15
                    Case 6
                        PuntosClubProducto = -18
                    Case 7
                        Select Case subgrup
                            Case 1
                                PuntosClubProducto = -8
                            Case 3
                                PuntosClubProducto = -2.5
                            Case Else
                                PuntosClubProducto = 0
                        End Select
                    Case Else
                        PuntosClubProducto = 0
                End Select

                
            Case #12/21/2022# To #10/17/2024#

                Select Case grup
                    Case 2
                        PuntosClubProducto = -12
                    Case 1, 3, 4, 8
                        PuntosClubProducto = -10
                    Case 24
                        PuntosClubProducto = -12
                    Case 14
                        PuntosClubProducto = -15
                    Case 6
                        PuntosClubProducto = -18
                    Case Else
                        PuntosClubProducto = 0
                End Select

                
             Case #10/16/2017# To #12/20/2022#
                Select Case Medida
                    Case "Kid"
                        PuntosClubProducto = -10
                    Case "Mediano"
                        PuntosClubProducto = -10
                    Case "Gigantona"
                        PuntosClubProducto = -10
                    Case "Bowl"
                        PuntosClubProducto = -15
                    Case "NV"
                        PuntosClubProducto = -7
                    Case Else
                        PuntosClubProducto = 0
                End Select
                
            Case Is < #10/16/2017#
                Select Case Medida
                    Case "Kid", "Mediano", "Gigantona"
                        PuntosClubProducto = -10
                    Case "Bowl"
                        PuntosClubProducto = -15
                    Case Else
                        PuntosClubProducto = 0
                End Select

        End Select
    
    Case 147 ' Agranda batidos x  p[untos
                
        PuntosClubProducto = -3

    Case 221 ' Agranda batidos x  p[untos
                
        PuntosClubProducto = -3
        
    Case 18  '2do a 20% Descuento
        Select Case fech
            Case Is < #10/16/2017#
                Select Case Medida
                    Case "Kid", "Mediano", "Gigantona", "Bowl"
                        PuntosClubProducto = 1
                    Case Else
                        PuntosClubProducto = 0
                End Select
            Case Else
                PuntosClubProducto = -1 ' actual
        End Select
    Case 21  '2do a 40%  Descuento
        Select Case fech
            Case Is < #10/16/2017#
                Select Case Medida
                    Case "Kid", "Mediano", "Gigantona", "Bowl"
                        PuntosClubProducto = 1
                    Case Else
                        PuntosClubProducto = 0
                End Select
            Case Else
                PuntosClubProducto = -1 ' actual
        End Select
    Case 48  '2do a 50% dscto
        PuntosClubProducto = -1
    Case 54  'Promociones 2do a 40% descuento
        PuntosClubProducto = -1
    Case 55  'Miercoles 3x2
        PuntosClubProducto = -2
    Case 60  'Segundo Mediano a C$39
        PuntosClubProducto = -1
    Case 61  'SSegundo Grande a C$47
        PuntosClubProducto = -1
    Case 68  '3x2 Batidos General
        PuntosClubProducto = -2
    Case 172  'Canjeado membresia x 1 punto
        PuntosClubProducto = -1
    Case 196  'Canjeado membresia x 1 punto
        PuntosClubProducto = -3
    Case 70  'S2do a 50% Dia de la Madre
        PuntosClubProducto = -1
    Case 204  '2do Batido a mitad de precio
        PuntosClubProducto = -1
    Case 90  'Lunes Waffle 10% dscto x compra batido, descuenta punto de batido
        PuntosClubProducto = -1
    Case 8, 6, 10, 3, 7, 31  'Promociones con Ratio 0, osea gratis
        PuntosClubProducto = 0
    
    Case 203
        PuntosClubProducto = -3
    Case Else ' Cualquier otra prmocion no aplica a puntos
        Select Case fech
            Case Is < #10/16/2017#
                Select Case Medida
                    Case "Kid", "Mediano", "Gigantona", "Bowl"
                        PuntosClubProducto = 1
                    Case Else
                        PuntosClubProducto = 0
                End Select
            Case Else
                PuntosClubProducto = 0
        End Select
End Select

Exit Function

Nulo:
PuntosClubProducto = 0

End Function



Function InicialesCliente(codc As Integer) As Long
' puntos iniciales de un codigo de club

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad total de puntos por el cliente usmando los inicales
miSQL = "SELECT ClientesClub.PuntosIniciales, ClientesClub.CodCliente" & _
" FROM ClientesClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((ClientesClub.CodCliente)=" & codc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
InicialesCliente = rst("PuntosIniciales")
rst.Close


Exit Function

Nulo:
InicialesCliente = 0


End Function

Function InicialesClienteLocal(codc As Long) As Long
' puntos iniciales de un codigo de club

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad total de puntos por el cliente usmando los inicales
miSQL = "SELECT ClientesClub.PuntosIniciales, ClientesClub.CodCliente" & _
" FROM ClientesClub" & _
" WHERE (((ClientesClub.CodCliente)=" & codc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
InicialesClienteLocal = rst("PuntosIniciales")
rst.Close


Exit Function

Nulo:
InicialesClienteLocal = 0


End Function

Function InicialesCliente2(codc As Integer) As Long
' puntos iniciales de un codigo de club

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad total de puntos por el cliente usmando los inicales
miSQL = "SELECT puntos_iniciales, membresia" & _
" FROM ClientesClubUnionExternoInterno" & _
" GROUP BY puntos_iniciales, membresia" & _
" HAVING (((membresia)=" & codc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
InicialesCliente2 = rst("puntos_iniciales")
rst.Close


Exit Function

Nulo:
InicialesCliente2 = 0
End Function


Function VentaClienteDia(Fecha As Date, club As Integer) As Long

'Ventas a codigo especifico de un dia
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Total de ventas a un codigo de cliente espcifico
miSQL = "SELECT NotaDePedido.CodCliente, NotaDePedido.Fecha, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Total FROM NotaDePedido GROUP BY NotaDePedido.CodCliente, NotaDePedido.Fecha HAVING (((NotaDePedido.CodCliente)=" & club & ") AND ((NotaDePedido.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentaClienteDia = rst("Total")
rst.Close

Exit Function

Nulo:
VentaClienteDia = 0

End Function

Function nombreCliente(club As Long) As String

'nombre del cliente club busqueda en mixed
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Solo nombre de liente club acorde a su codigo
miSQL = "SELECT ClientesClub.CodCliente, ClientesClub.Nombre FROM ClientesClub" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((ClientesClub.CodCliente)=" & club & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreCliente = rst("Nombre")
rst.Close

Exit Function

Nulo:
nombreCliente = ""

End Function

Function iniciocliente(club As Integer) As Date

'fecha de inscripcion de cliente
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'La fecha en que fue inscrito al sistema
miSQL = "SELECT ClientesClub.CodCliente, ClientesClub.[Fecha de Inscripcion] FROM ClientesClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb' WHERE (((ClientesClub.CodCliente)=" & club & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
iniciocliente = rst("Fecha de Inscripcion")
rst.Close

Exit Function

Nulo:
iniciocliente = ""

End Function

Function totalbatidogratiscumple(club As Integer, cano As Integer) As Double
'cuantos productos canjeados por cumpleaños en el año
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'La fecha en que fue inscrito al sistema
miSQL = "SELECT DetallePedidoClub.CodPromocion, Year([DetallePedidoClub]![Fecha]) AS cano," & _
" DetallePedidoClub.CodCliente, Sum(DetallePedidoClub.Cantidad) AS SumaDeCantidad" & _
" FROM DetallePedidoClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY DetallePedidoClub.CodPromocion, Year([DetallePedidoClub]![Fecha]), DetallePedidoClub.CodCliente" & _
" HAVING (((DetallePedidoClub.CodPromocion)=8) AND ((Year([DetallePedidoClub]![Fecha]))=" & cano & ")" & _
" AND ((DetallePedidoClub.CodCliente)=" & club & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
totalbatidogratiscumple = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
totalbatidogratiscumple = 0

End Function

Function fechacumpleclub(clubi As Long) As Date
'fecha de cumpleaños de cliente club
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT ClientesClub.CodCliente, ClientesClub.Cumpleanos" & _
" FROM ClientesClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((ClientesClub.CodCliente)=" & clubi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
fechacumpleclub = rst("Cumpleanos")
rst.Close

Exit Function

Nulo:
fechacumpleclub = 0

End Function

Function ultimacompracliente(clubx As Long) As Date
' ultima comrpa hecha por lciente club , compra con monto diferente de cero

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad total de puntos por el cliente usmando los inicales
miSQL = "SELECT NotadePedido.CodCliente, NotadePedido.TotalGuardado, NotadePedido.Fecha" & _
" FROM NotadePedido IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((NotadePedido.CodCliente) = " & clubx & ") And ((NotadePedido.TotalGuardado) <> 0) And ((NotadePedido.Fecha) <> #" & Date & "#))" & _
" ORDER BY NotadePedido.Fecha DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimacompracliente = rst("Fecha")
rst.Close


Exit Function

Nulo:
ultimacompracliente = 0


End Function

Function ultimoclientedeliveryregistrado() As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT ClientesDelivery.CodClientesDelivery FROM ClientesDelivery ORDER BY ClientesDelivery.CodClientesDelivery DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimoclientedeliveryregistrado = rst("CodClientesDelivery")
rst.Close
Exit Function

Nulo:
ultimoclientedeliveryregistrado = 0
rst.Close
End Function

Sub CargarVentasInternoExternoClienteHaciaTabla(ByVal pCodCliente As Long)
'juntar puntos descargados de web (externos) con puntos de DB local access en una tabla VentasGlobalesFiltradoClienteInternoExterno

    Dim db As DAO.Database
    Dim sql As String
    Dim tdf As DAO.TableDef
    Dim existeTabla As Boolean

    Set db = CurrentDb
    existeTabla = False

    ' 1?? Verificar si la tabla existe
    For Each tdf In db.TableDefs
        If tdf.Name = "VentasGlobalesFiltradoClienteInternoExterno" Then
            existeTabla = True
            Exit For
        End If
    Next tdf

    ' 2?? Crear tabla si no existe
    If Not existeTabla Then
        db.Execute _
            "CREATE TABLE VentasGlobalesFiltradoClienteInternoExterno (" & _
            "Anulado YESNO, " & _
            "Fecha DATETIME, " & _
            "Hora DATETIME, " & _
            "CodPedido LONG, " & _
            "CodCliente LONG, " & _
            "DBBatidos_Nombre TEXT(255), " & _
            "Medida TEXT(50), " & _
            "Cantidad DOUBLE, " & _
            "CodigoPromocion LONG, " & _
            "local LONG, " & _
            "Puntos DOUBLE, " & _
            "PuntosLinea DOUBLE, " & _
            "CodProducto TEXT(255), " & _
            "Sucursal_Nombre TEXT(255), " & _
            "Origen TEXT(20)" & _
            ")", dbFailOnError
    End If

    ' 3?? Limpiar tabla
    db.Execute _
        "DELETE FROM VentasGlobalesFiltradoClienteInternoExterno", _
        dbFailOnError

    ' 4?? INSERT registros INTERNOS
    sql = "INSERT INTO VentasGlobalesFiltradoClienteInternoExterno (" & _
          "Anulado, Fecha, Hora, CodPedido, CodCliente, DBBatidos_Nombre, Medida, " & _
          "Cantidad, CodigoPromocion, local, Puntos, PuntosLinea, CodProducto, " & _
          "Sucursal_Nombre, Origen) " & _
          "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, NotaDePedido.Hora, " & _
          "NotaDePedido.CodPedido, NotaDePedido.CodCliente, " & _
          "DBBatidos.Nombre, DBBatidos.Medida, SubPedido.Cantidad, " & _
          "SubPedido.CodPromocion, codigolocal(), SubPedido.Puntos, " & _
          "(SubPedido.Cantidad * SubPedido.Puntos), SubPedido.CodBatido, " & _
          "nombrelocalglobal(codigolocal()), 'Interno' " & _
          "FROM (SubPedido INNER JOIN DBBatidos " & _
          "ON SubPedido.CodBatido = DBBatidos.CodBatido) " & _
          "INNER JOIN NotaDePedido " & _
          "ON SubPedido.CodPedido = NotaDePedido.CodPedido " & _
          "WHERE NotaDePedido.CodCliente = " & pCodCliente

    db.Execute sql, dbFailOnError

    ' 5?? INSERT registros EXTERNOS
    sql = "INSERT INTO VentasGlobalesFiltradoClienteInternoExterno (" & _
          "Anulado, Fecha, Hora, CodPedido, CodCliente, DBBatidos_Nombre, Medida, " & _
          "Cantidad, CodigoPromocion, local, Puntos, PuntosLinea, CodProducto, " & _
          "Sucursal_Nombre, Origen) " & _
          "SELECT Anulado, Fecha, Hora, CodPedido, CodCliente, DBBatidos_Nombre, " & _
          "Medida, Cantidad, CodigoPromocion, local, Puntos, " & _
          "(Cantidad * Puntos), CodProducto, Sucursal_Nombre, 'Externo' " & _
          "FROM VentasGlobalesAccessCSVFiltradoCliente"

    db.Execute sql, dbFailOnError

    Set db = Nothing

End Sub


Function puntosclubsolocanjeofactura(codigopedido As Long) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Sum(SubPedido.Puntos * Nz(SubPedido.Cantidad, 0)) AS SumaDePuntos" & _
" FROM SubPedido" & _
" WHERE SubPedido.CodPromocion = 22 AND SubPedido.CodPedido = " & codigopedido
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
If Not rst.EOF Then
    puntosclubsolocanjeofactura = Nz(rst("SumaDePuntos"), 0)
Else
    puntosclubsolocanjeofactura = 0
End If
rst.Close
Set rst = Nothing
Exit Function

Nulo:
puntosclubsolocanjeofactura = 0
If Not rst Is Nothing Then
    rst.Close
    Set rst = Nothing
End If
End Function



