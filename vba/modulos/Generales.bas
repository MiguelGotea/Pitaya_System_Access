' ==========================================================
' Modulo  : Generales
' Tipo    : 1
' Lineas  : 1689
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database
Function ListarTablasVinculadasAgrupadasCompacto() As String
    On Error GoTo ManejarError
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim dict As Object
    Dim listaTablas As String
    Dim ruta As String
    Dim clave As Variant
    Dim contadorTotal As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set db = CurrentDb()
    contadorTotal = 0
    
    ' Agrupar tablas por ruta
    For Each tdf In db.TableDefs
        If InStr(tdf.Connect, ".accdb") > 0 And Len(tdf.Connect) > 0 Then
            ruta = ObtenerRutaArchivo(tdf.Connect)
            contadorTotal = contadorTotal + 1
            
            If Not dict.Exists(ruta) Then
                dict.Add ruta, ""
            End If
            
            ' Concatenar nombres de tablas separados por coma
            If dict(ruta) = "" Then
                dict(ruta) = tdf.Name
            Else
                dict(ruta) = dict(ruta) & ", " & tdf.Name
            End If
        End If
    Next tdf
    
    ' Construir mensaje compacto
    listaTablas = ""
    
    If dict.Count = 0 Then
        listaTablas = listaTablas & "No se encontraron tablas vinculadas."
    Else
        For Each clave In dict.Keys
            listaTablas = listaTablas & "Vinculados desde " & clave & ":" & vbCrLf
            listaTablas = listaTablas & dict(clave) & vbCrLf
            listaTablas = listaTablas & "(" & (Len(dict(clave)) - Len(Replace(dict(clave), ",", "")) + 1) & " tablas)" & vbCrLf & vbCrLf
        Next clave
        
        listaTablas = listaTablas & "TOTAL: " & contadorTotal & " tablas en " & dict.Count & " rutas"
    End If
    
    ListarTablasVinculadasAgrupadasCompacto = listaTablas
    
LimpiarObjetos:
        Set tdf = Nothing
        Set db = Nothing
        Set dict = Nothing
    Exit Function
    
ManejarError:
    ListarTablasVinculadasAgrupadasCompacto = "Error: " & Err.Description
    Resume LimpiarObjetos
End Function

Function ObtenerRutaArchivo(connectString As String) As String
    Dim ruta As String
    
    If InStr(connectString, "DATABASE=") > 0 Then
        ruta = Mid(connectString, InStr(connectString, "DATABASE=") + 9)
        ' Limpiar la ruta si tiene comillas
        If Left(ruta, 1) = ";" Then ruta = Mid(ruta, 2)
        ObtenerRutaArchivo = Replace(ruta, ";", "")
        ObtenerRutaArchivo = Split(ObtenerRutaArchivo, "Google Drive BP\")(1)
    Else
        ObtenerRutaArchivo = connectString
    End If
End Function

Sub multiplebeep(ca As Integer)
Dim con As Integer

For con = 1 To ca
    DoCmd.Beep
    Sleep 1000
Next
End Sub

Function AbrirCalculadora()
DoCmd.SetWarnings False
Application.FollowHyperlink "C:\Windows\system32\calc.exe"
DoCmd.SetWarnings True
End Function
Function TarjetaDisponble() As Long

'Siguinte tarjeta disponible

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la ultima tarjeta vendida
miSQL = "SELECT Max(ClientesClub.CodCliente) AS MáxDeCodCliente FROM ClientesClub"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TarjetaDisponble = rst("MáxDeCodCliente") + 1
rst.Close
     
Exit Function

Nulo:
TarjetaDisponble = 0

End Function


Function redondear_mas(num As Double) As Long

redondear_mas = Int(num) + IIf(num - Int(num) > 0, 1, 0)

End Function

Function si_negativo_cero(num As Double) As Double

If num < 0 Then
si_negativo_cero = 0
Else
si_negativo_cero = num
End If


End Function

Function ProovedorActual(coti As Integer) As Long
'Provvedor con mayor compra el ultimo mes de cierta cotizacion

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Histoprial de compras de ultimo mes, se elige proovedor con mayores comporas en cantidad
miSQL = "SELECT Sum(Compras.Cantidad) AS SumaDeCantidad, Compras.CodCotizacion, numerosemana([Compras]![Fecha])>numerosemana(#" & Date & "#)-5 AS Expr1, Compras.CodProveedor AS Proovedor FROM Compras GROUP BY Compras.CodCotizacion, numerosemana([Compras]![Fecha])>numerosemana(#" & Date & "#)-5, Compras.CodProveedor HAVING (((Compras.CodCotizacion) = " & coti & ") And ((numerosemana([Compras]![Fecha]) > numerosemana(#" & Date & "#) - 5) = True)) ORDER BY Sum(Compras.Cantidad) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ProovedorActual = rst("Proovedor")
rst.Close

Exit Function

Nulo:
ProovedorActual = 38 'proovedor nn sin regisdtro

End Function
Function ProovedorPrincipal(codip As Integer) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'
miSQL = "SELECT Cotizaciones.CodCotizacion, DBIngredientes.ProovedorPrincipal" & _
" FROM Cotizaciones INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((Cotizaciones.CodCotizacion)=" & codip & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ProovedorPrincipal = rst("ProovedorPrincipal")
rst.Close
     
Exit Function

Nulo:
ProovedorPrincipal = 0

End Function
Function NombreProovedor(codip As Integer) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'
miSQL = "SELECT Proovedores.Nombre, Proovedores.CodProovedor" & _
" FROM Proovedores" & _
" WHERE (((Proovedores.CodProovedor)=" & codip & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
NombreProovedor = rst("Nombre")
rst.Close
     
Exit Function

Nulo:
NombreProovedor = ""

End Function

Function ConsumoEmpaque(pedido As Double, Medida As String, Tipo As Integer) As Double

'fecha: fecha de consulta
'medida: tipo de consulta kid mediano gigantona o bowl
'tipo: tipo de dato True: Vidrio, False: plastico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de material de empaque desgun consulta de un subpedido especifico
miSQL = "SELECT DBBatidos.Medida, NotaDePedido.Modalidad, SubPedido.CodSubPedido, SubPedido.Cantidad FROM DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) ON DBBatidos.CodBatido = SubPedido.CodBatido WHERE (((DBBatidos.Medida)='" & Medida & "') AND ((NotaDePedido.Modalidad)=" & Tipo & ") AND ((SubPedido.CodSubPedido)=" & pedido & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoEmpaque = rst("Cantidad")
rst.Close

Exit Function

Nulo:
ConsumoEmpaque = 0

End Function

Function clavesistema(formu As String) As Long

Dim clave As String

clave = InputBox("Ingresar clave:", "Ingreso Seguridad")

If clave = "miguel" Then

Else
MsgBox "Clave Incorrecta"

DoCmd.Close acForm, formu

End If

End Function
Function numerosemana(fech As Date) As Long

Select Case Year(fech)
    Case 2016 'Inicio de semanas 1 el 1/1/2016
        numerosemana = DatePart("WW", fech, vbMonday)
    Case 2017
        numerosemana = DatePart("WW", fech, vbMonday) + 52  '52
    Case 2018
        numerosemana = DatePart("WW", fech, vbMonday) + 105 '53
    Case 2019
        numerosemana = DatePart("WW", fech, vbMonday) + 157 '52
    Case 2020
        numerosemana = DatePart("WW", fech, vbMonday) + 209 '52
    Case 2021
        numerosemana = DatePart("WW", fech, vbMonday) + 261 '52
    Case 2022
        numerosemana = DatePart("WW", fech, vbMonday) + 313 '52
    Case 2023
        numerosemana = DatePart("WW", fech, vbMonday) + 365 '52
    Case 2024
        numerosemana = DatePart("WW", fech, vbMonday) + 418 '53
    Case 2025
        numerosemana = DatePart("WW", fech, vbMonday) + 470 '52
    Case 2026
        numerosemana = DatePart("WW", fech, vbMonday) + 522 '52
    Case 2027
        numerosemana = DatePart("WW", fech, vbMonday) + 574 '52
    Case 2028
        numerosemana = DatePart("WW", fech, vbMonday) + 626 '52
    Case 2029
        numerosemana = DatePart("WW", fech, vbMonday) + 679 '53
    Case 2030
        numerosemana = DatePart("WW", fech, vbMonday) + 731 '52
    Case 2031
        numerosemana = DatePart("WW", fech, vbMonday) + 783 '52
    Case 2032
        numerosemana = DatePart("WW", fech, vbMonday) + 835 '52
    Case 2033
        numerosemana = DatePart("WW", fech, vbMonday) + 887 '52
    Case 2034
        numerosemana = DatePart("WW", fech, vbMonday) + 939 '52
    Case 2035
        numerosemana = DatePart("WW", fech, vbMonday) + 992 '52
    Case 2036
        numerosemana = DatePart("WW", fech, vbMonday) + 1044 '52
    Case 2037
        numerosemana = DatePart("WW", fech, vbMonday) + 1096 '52
    Case 2038
        numerosemana = DatePart("WW", fech, vbMonday) + 1148 '52
    Case 2039
        numerosemana = DatePart("WW", fech, vbMonday) + 1200 '52
    Case 2040
        numerosemana = DatePart("WW", fech, vbMonday) + 1252 '52
    Case 2041
        numerosemana = DatePart("WW", fech, vbMonday) + 1305 '52
    Case 2042
        numerosemana = DatePart("WW", fech, vbMonday) + 1357 '52
    Case 2043
        numerosemana = DatePart("WW", fech, vbMonday) + 1409 '52
    Case Else
        numerosemana = 0
End Select


End Function
Function diasmes(fech As Date) As Long

diasmes = Day(DateSerial(Year(fech), Month(fech) + 1, 0))

End Function


Function TipoDeCambio() As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Contador Inicial de selladora
miSQL = "SELECT EstadoInicial.Fecha, EstadoInicial.[TipoCambio$_C$] FROM EstadoInicial WHERE (((EstadoInicial.Fecha)=Date()))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TipoDeCambio = rst("[TipoCambio$_C$]")
rst.Close

Exit Function

Nulo:
'MsgBox "No Hay Datos Iniciales Ingresados"
TipoDeCambio = 0

End Function

Function FactorDeUso(ingre As String, pedido As Long) As Long
FactorDeUso = 1
'ing ingrediente
'pedido codigo de nota de pedido

'Dim tipoe As String
'tipoe = usovidrio(pedido)

'Select Case ingre
'Case "P001", "P002", "P003", "P004", "P006", "P007", "P008", "P009", "P011", "P015", "P016", "P017"
'    FactorDeUso = IIf(tipoe = "Plastico", 1, 0)
'Case "P005"
'    FactorDeUso = IIf(tipoe = "Plastico", 0, 1)
'Case Else
'    FactorDeUso = 1
'End Select

End Function

Function FactorDeUsoMejorado(ingre As String, empaq As Integer, estadovidrioantiguo As Integer, fechapedido As Date) As Integer

Select Case fechapedido
    Case Is < #1/1/2024#
        Select Case ingre
            Case "P001", "P002", "P003", "P004", "P006", "P007", "P008", "P009", "P011", "P015", "P016", "P017"
                FactorDeUsoMejorado = IIf(estadovidrioantiguo = 0, 1, 0)
            Case "P005"
                FactorDeUsoMejorado = IIf(estadovidrioantiguo = 0, 0, 1)
            Case Else
                FactorDeUsoMejorado = 1
        End Select
    
    Case Else
        Select Case ingre
            Case "P016", "P017", "P011", "P026"
                FactorDeUsoMejorado = IIf(empaq = 0, 0, 1)
            Case Else
                FactorDeUsoMejorado = 1
        End Select
End Select

End Function
Function CantidadIngrediente(bat As String, ingre As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la cantidad del ingrediente especifico en la receta especifica
miSQL = "SELECT SubReceta.CodBatido, SubReceta.CodIngrediente, Sum(SubReceta.Cantidad) AS SumaDeCantidad" & _
" FROM SubReceta" & _
" GROUP BY SubReceta.CodBatido, SubReceta.CodIngrediente" & _
" HAVING (((SubReceta.CodBatido)='" & bat & "') AND ((SubReceta.CodIngrediente)='" & ingre & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadIngrediente = rst("SumaDeCantidad")
rst.Close
     
Exit Function

Nulo:
CantidadIngrediente = 0

End Function
Function FactorTiempoProduccion(codsubped As Long) As Long
' Factor de tiempod e prduccion de batido 1 y de bowl 3 codigo de subpedido

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'CHallar la cantidad por el factor
miSQL = "SELECT Grupos.Tipo AS Tipo, SubPedido.CodSubPedido, SubPedido.Cantidad AS Total FROM SubPedido INNER JOIN (Grupos INNER JOIN DBBatidos ON Grupos.CodGrupo = DBBatidos.CodGrupo) ON SubPedido.CodBatido = DBBatidos.CodBatido WHERE (((Grupos.Tipo)='Batido' Or (Grupos.Tipo)='Bowl') AND ((SubPedido.CodSubPedido)=" & codsubped & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If rst("Tipo") = "Batido" Then
    FactorTiempoProduccion = rst("Total")
    rst.Close
Else
    FactorTiempoProduccion = rst("Total") * 3
    rst.Close
End If

Exit Function

Nulo:
FactorTiempoProduccion = 0

End Function

Function ultimodomingomes(Mes As Integer, ano As Integer) As Date

ultimodia = diasmes(DateSerial(ano, Mes, 1))
For D = ultimodia To 1 Step -1
If Weekday(DateSerial(ano, Mes, D), vbMonday) = 7 Then
ultimodomingomes = DateSerial(ano, Mes, D)
Exit Function
End If

Next D

End Function

Function nombrelocal() As String
On Error GoTo Nulo
nombrelocal = DLookup("[Nombre]", "DatosSistema")
Exit Function

Nulo:
nombrelocal = "Pitaya"
End Function

Function nombrelocalglobal(codi As Integer) As String
On Error GoTo Nulo
nombrelocalglobal = DLookup("[Nombre]", "[StatusSucursales]", "[CodLocal]=" & codi)
Exit Function

Nulo:
nombrelocalglobal = "Pitaya"
End Function

Function codigoLocal() As Long
On Error GoTo Nulo
codigoLocal = DLookup("[CodSistema]", "DatosSistema")
Exit Function

Nulo:
codigoLocal = 0
End Function

Function clavewifi() As String

clavewifi = DLookup("[clavewifi]", "DatosSistema")

End Function
Function ciudadsistema() As String
On Error GoTo Nulo
ciudadsistema = DLookup("[Ciudad]", "DatosSistema")
Exit Function

Nulo:
ciudadsistema = "General"
End Function

Function fraccion(num As Double) As String
Dim numerador As Integer
Dim denominador As Integer
Dim entero As Integer
Dim decima As Double

entero = Int(num)
decima = num - entero

If decima > 0 Then ' si no hay fraccion no aplica
    numerador = decima * 8
    denominador = 8

Select Case Round(decima, 2)
        Case 0.5
            numerador = 1
            denominador = 2
        Case 0.33
            numerador = 1
            denominador = 3
        Case 0.25
            numerador = 1
            denominador = 4
        Case 0.17
            numerador = 1
            denominador = 6
        Case 0.13
            numerador = 1
            denominador = 8
        Case 0.67
            numerador = 2
            denominador = 3
        Case 0.75
            numerador = 3
            denominador = 4
        Case 0.83
            numerador = 5
            denominador = 6
    End Select
End If

'Definir fraccion
If decima = 0 Then
    fraccion = entero
ElseIf entero = 0 Then
    fraccion = numerador & "/" & denominador
Else
    fraccion = entero & "+" & numerador & "/" & denominador
End If

End Function

Function mcm(inge As String, por As Integer, Tipo As String, Cond As Integer) As Double
'COnd:
'por: medida de receta
'0: cantidad o tamano de porcion
'1: factor de mutipicador
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TOP 1 DBBatidos.Vigencia, SubReceta.CodIngrediente, SubReceta.Cantidad, SubReceta.codporcion" & _
" FROM DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" WHERE (((SubReceta.CodIngrediente)='" & inge & "') AND ((SubReceta.Cantidad)=" & por & "))" '((DBBatidos.Vigencia)<>0) AND
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Cond = 0 Then
    mcm = DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion"))
Else
    mcm = por / DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & rst("codporcion"))
End If
rst.Close

Exit Function

Nulo:
mcm = 0

End Function

Function IngresarPesoPorcionar() As Long
IngresarPesoPorcionar = InputBox("Ingresar peso a porcionar", "INGRESAR PESO")
End Function

Function MarcaProducto(coti As Integer) As String
On Error Resume Next
MarcaProducto = DLookup("[Marca]", "Cotizaciones", "[CodCotizacion] = " & coti)
End Function

Function esporcionable(ING As String) As Long
'-1:verdadero
'0:Falso
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Saber si existen porciones en la listade cotizacion de un ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente, Cotizaciones.Subproducto" & _
" From Cotizaciones" & _
" GROUP BY Cotizaciones.CodIngrediente, Cotizaciones.Subproducto" & _
" HAVING (Cotizaciones.CodIngrediente='" & ING & "' AND Cotizaciones.Subproducto<>0)"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
esporcionable = rst("Subproducto")
rst.Close
     
Exit Function

Nulo:
esporcionable = 0

End Function

Function busquedacotizacionporcion(medi As Integer, ING As String) As Long
'Cantidad cotizaciones de porcion usadas en una semana

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.CodIngrediente, Cotizaciones.Subproducto, Cotizaciones.Conversion, Cotizaciones.CodCotizacion From Cotizaciones WHERE (((Cotizaciones.CodIngrediente)='" & ING & "') AND ((Cotizaciones.Subproducto)<>0) AND ((Cotizaciones.Conversion)=" & medi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
busquedacotizacionporcion = rst("CodCotizacion")
rst.Close

Exit Function

Nulo:
busquedacotizacionporcion = 0

End Function

Function cantidadxmdidaxtipo(ING As String) As Long
'sumatoriade cantidad de pedidos de una mediad por el peso de la porcion , la suma total por tipo degrupo de porcion

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Integer
Dim Cant As Double
suma = 0

miSQL = "SELECT Grupos.Tipo, mcm([DBIngredientes]![CodIngrediente],[SubReceta]![Cantidad],[Grupos]![Tipo],0) AS Cantidad, SubReceta.CodIngrediente, DBBatidos.Vigencia FROM Grupos INNER JOIN (DBBatidos INNER JOIN (DBIngredientes INNER JOIN SubReceta ON DBIngredientes.CodIngrediente = SubReceta.CodIngrediente) ON DBBatidos.CodBatido = SubReceta.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo GROUP BY Grupos.Tipo, mcm([DBIngredientes]![CodIngrediente],[SubReceta]![Cantidad],[Grupos]![Tipo],0), SubReceta.CodIngrediente, DBBatidos.Vigencia HAVING (((SubReceta.CodIngrediente)='" & ING & "') AND ((DBBatidos.Vigencia)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
Cant = rst("Cantidad")
suma = suma + Cant * ConsumoPorPorciones(ING, Cant, rst("Tipo"))
rst.MoveNext
Loop
rst.Close
cantidadxmdidaxtipo = suma
Exit Function

Nulo:
cantidadxmdidaxtipo = 0

End Function

Function esporcionablecotizacion(coti As Integer) As Long
'consultar si es porcionable una unidad de cotizacion
'-1: porcion
'0: no porcion

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.Subproducto, Cotizaciones.CodCotizacion FROM Cotizaciones WHERE (((Cotizaciones.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
esporcionablecotizacion = rst("Subproducto")
Exit Function

Nulo:
esporcionablecotizacion = 0

End Function

Function AgregarPropAp(strName As String, varType As Variant, varValue As Variant) As Long
'FUncion para agregar tituloe icono de manera directa de macro
Dim dbs As Object, prp As Variant
Const conPropNotFoundError = 3270

Set dbs = CurrentDb
On Error GoTo AddProp_Err
dbs.Properties(strName) = varValue
AddAppProperty = True
 
AddProp_Bye:
    Exit Function
 
AddProp_Err:
If Err = conPropNotFoundError Then
    Set prp = dbs.CreateProperty(strName, varType, varValue)
    dbs.Properties.Append prp
    Resume
Else
    AddAppProperty = False
    Resume AddProp_Bye
    End If
End Function

Function AgregarTitulo()
Dim intX As Integer
Const dbText As Long = 10
intX = AgregarPropAp("AppTitle", dbText, "Batidos Pitaya " & codigoLocal() & " " & ciudadsistema()) 'Nombre de sistema
intX = AgregarPropAp("AppIcon", dbText, "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Logo\Logo.ico") 'Ruta de icono
Application.RefreshTitleBar 'Refresco el barra de título de la aplicación
End Function

Function ReiniciarSistema()
If MsgBox("Desea reiniciar el sistema?", vbYesNo, "Acceso Rapido") = vbYes Then
    Application.FollowHyperlink "C:\Users\" & NombreSistema() & "\Desktop\Sistema\reiniciar_sistema.vbs"
End If
End Function

Public Function DescripcionFormulario(ByVal ItemName As String) As Variant
Dim result As Variant
On Error GoTo exitfunction
result = CurrentDb.Containers("Forms").Documents(ItemName).Properties("Description")
DescripcionFormulario = result
Exit Function

exitfunction:
DescripcionFormulario = Null
End Function

Function mensajesdialocal(loc As Integer) As Long
'mensajes totdakes ekvados de un local especifico en el dia

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DateValue(x.Fecha) AS Expr1, x.Emisor, x.Receptor, Count(x.Emisor) AS CuentaDeEmisor" & _
" FROM (SELECT *," & loc & " as Emisor FROM ChatPitaya" & _
" IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & loc & "_DB.accdb'" & _
" WHERE Receptor = codigolocal()" & _
" UNION ALL SELECT *,codigolocal() as Emisor FROM ChatPitaya" & _
" IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & codigoLocal() & "_DB.accdb'" & _
" WHERE Receptor = " & loc & ")x" & _
" GROUP BY DateValue(x.Fecha), x.Emisor, x.Receptor" & _
" HAVING DateValue(x.Fecha)=Date() AND x.Emisor=" & loc & " AND x.Receptor= codigolocal()"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
mensajesdialocal = rst("CuentaDeEmisor")

Exit Function

Nulo:
mensajesdialocal = 0

End Function
Function nombreproductoventa(bati As String) As String
'

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBBatidos.Nombre, DBBatidos.Medida, DBBatidos.CodBatido FROM DBBatidos" & _
" WHERE (((DBBatidos.CodBatido)='" & bati & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreproductoventa = rst("Nombre") & " " & IIf(rst("Medida") = "NV", "", rst("Medida"))
Exit Function

Nulo:
nombreproductoventa = "sin nombre"

End Function
Function unidadproductocoti(coti As Integer) As String
' [DBIngredientes]![NombreSinProcesar] & " " & [Cotizaciones]![Marca] & " " & [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [Cotizaciones]![Marca] & ' ' & [Cotizaciones]![Linea] & ' ' & [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS nombre, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (Cotizaciones.CodCotizacion=" & coti & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
unidadproductocoti = rst("nombre")
Exit Function

Nulo:
unidadproductocoti = "sin presentacion"

End Function
Function nombreproductocoti(coti As Integer) As String
' [DBIngredientes]![NombreSinProcesar] & " " & [Cotizaciones]![Marca] & " " & [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [DBIngredientes]![Nombre] & ' ' & IIf(IsNull([Cotizaciones]![Marca]),'',[Cotizaciones]![Marca]) & ' ' & IIf(IsNull([Cotizaciones]![Linea]),'',[Cotizaciones]![Linea]) & ' ' & IIf(IsNull([Cotizaciones]![Unidad]),'',[Cotizaciones]![Unidad]) & ' ' & IIf(IsNull([Cotizaciones]![Capacidad]),'',[Cotizaciones]![Capacidad]) AS nombre, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (((Cotizaciones.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreproductocoti = rst("nombre")
Exit Function

Nulo:
nombreproductocoti = "sin nombre"

End Function


Function nombreproductocotiprocesado(coti As Integer) As String
' [DBIngredientes]![Nombre] & " " & [Cotizaciones]![Marca] & " " & [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT IIf([Cotizaciones]![Conversion]=0,[DBIngredientes]![NombreSinProcesar],[DBIngredientes]![NombreProcesado]) & ' ' & IIf(IsNull([Cotizaciones]![Marca]),'',[Cotizaciones]![Marca]) & ' ' & IIf(IsNull([Cotizaciones]![Linea]),'',[Cotizaciones]![Linea]) & ' ' & IIf(IsNull([Cotizaciones]![Unidad]),'',[Cotizaciones]![Unidad]) & ' ' & IIf(IsNull([Cotizaciones]![Capacidad]),'',[Cotizaciones]![Capacidad]) AS nombre, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (Cotizaciones.CodCotizacion=" & coti & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreproductocotiprocesado = rst("nombre")
Exit Function

Nulo:
nombreproductocotiprocesado = "sin nombre"

End Function
Function nombreproductocotiprocesadosolounidad(coti As Integer) As String
' [DBIngredientes]![Nombre] & " " & [Cotizaciones]![Marca] & " " & [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT IIf(IsNull([Cotizaciones]![Unidad]),'',[Cotizaciones]![Unidad]) & ' ' & IIf(IsNull([Cotizaciones]![Capacidad]),'',[Cotizaciones]![Capacidad]) AS nombre, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (Cotizaciones.CodCotizacion=" & coti & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreproductocotiprocesadosolounidad = rst("nombre")
Exit Function

Nulo:
nombreproductocotiprocesadosolounidad = "sin nombre"

End Function
Function nombreproductocotiprocesadosinunidad(coti As Integer) As String
' [DBIngredientes]![Nombre] & " " & [Cotizaciones]![Marca] & " " & [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT IIf([Cotizaciones]![Conversion]=0,[DBIngredientes]![NombreSinProcesar],[DBIngredientes]![NombreProcesado]) & ' ' & IIf(IsNull([Cotizaciones]![Marca]),'',[Cotizaciones]![Marca]) & ' ' & IIf(IsNull([Cotizaciones]![Linea]),'',[Cotizaciones]![Linea]) AS nombre, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (Cotizaciones.CodCotizacion=" & coti & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreproductocotiprocesadosinunidad = rst("nombre")
Exit Function

Nulo:
nombreproductocotiprocesadosinunidad = "sin nombre"

End Function
Function nombreproductoporcionreceta(coti As Integer) As String
' [DBIngredientes]![NombreSinProcesar] & " " & [Cotizaciones]![Marca] & " " & [Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [DBIngredientes]![Nombre] & ' (' & [Cotizaciones]![Capacidad] & ')' AS nombre, Cotizaciones.CodCotizacion" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (Cotizaciones.CodCotizacion=" & coti & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreproductoporcionreceta = rst("nombre")
Exit Function

Nulo:
nombreproductoporcionreceta = "sin nombre"

End Function
Function temporalarrayunion(Tabla As String) As String

Dim totallocales As Integer

totallocales = CDbl(consultasql("SELECT Count(Codigo) AS CuentaDeCodigo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
"", "CuentaDeCodigo"))

temporalarrayunion = ""

For I = 1 To totallocales
    temporalarrayunion = temporalarrayunion & IIf(I = 1, "", "UNION ALL ") & "SELECT * FROM " & Tabla & " IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & I & "_DB.accdb' "
Next I

End Function

Function consultasql(miSQL As String, Valor As String) As String
On Error GoTo Nulo

Dim rst As DAO.Recordset
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consultasql = CStr(rst(Valor))
rst.Close


Exit Function

Nulo:
consultasql = ""

End Function

Function MaxMinSemana(sem As Integer, Cond As Integer) As Date
'fecha maxima (1) y minima (0) de una semana especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String


'Calcula max fecha de una seman especifica
miSQL = "SELECT TOP 1 numerosemana([FechaSistema]![Dates]) AS semana, FechaSistema.Dates" & _
" FROM fechasistema" & _
" WHERE (((numerosemana([fechasistema]![Dates])) = " & sem & "))" & _
" ORDER BY FechaSistema.Dates " & IIf(Cond = 1, "DESC", "ASC")
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MaxMinSemana = rst("Dates")
rst.Close



Exit Function

Nulo:
MaxMinSemana = 0

End Function

Function FechaDeNumeroDiaSemana(sem As Integer, diax As Integer) As Date
'fecha de una semana segun el numero de dia enviado del 1 al 7

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
'
miSQL = "SELECT numerosemana([FechaSistema]![Dates]) AS semana, Weekday([FechaSistema]![Dates],2) AS secuencia, fechasistema.Dates" & _
" FROM fechasistema" & _
" WHERE (((numerosemana([fechasistema]![Dates])) = " & sem & ") And ((Weekday([fechasistema]![Dates], 2)) = " & diax & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FechaDeNumeroDiaSemana = rst("Dates")
rst.Close

Exit Function

Nulo:
FechaDeNumeroDiaSemana = 0
rst.Close
End Function

Function PorcionDeCotizacion(coti As Long) As Long
'buscar la porcion subproducto<>0, de cotizacion Almacen Principal
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim inge As String
Dim conv As Integer
Dim line As String

inge = DLookup("[CodIngrediente]", "Cotizaciones", "[CodCotizacion] = " & coti)
conv = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti)
line = IIf(IsNull(DLookup("[Linea]", "Cotizaciones", "[CodCotizacion] = " & coti)), "", DLookup("[Linea]", "Cotizaciones", "[CodCotizacion] = " & coti))

miSQL = "SELECT Subproducto, CodIngrediente, Conversion, CodCotizacion, Linea" & _
" FROM Cotizaciones" & _
" WHERE (Subproducto<>0" & _
IIf(line = "", "", " AND Linea='" & line & "'") & _
" AND CodIngrediente='" & inge & "' AND Conversion=" & conv & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PorcionDeCotizacion = rst("CodCotizacion")
rst.Close

Exit Function

Nulo:
PorcionDeCotizacion = 0

End Function

Function PorcionGlobalDePorcion(coti As Long) As Long
'buscar la porcion marca Almacen Globa, de cotizacion porcion subprod<>0
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim inge As String
Dim conv As Double
Dim line As String

inge = DLookup("[CodIngrediente]", "Cotizaciones", "[CodCotizacion] = " & coti)
conv = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti)
line = IIf(IsNull(DLookup("[Linea]", "Cotizaciones", "[CodCotizacion] = " & coti)), "", DLookup("[Linea]", "Cotizaciones", "[CodCotizacion] = " & coti))

miSQL = "SELECT Cotizaciones.Marca, Cotizaciones.Linea, Cotizaciones.Conversion, Cotizaciones.CodIngrediente, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones" & _
" WHERE (Cotizaciones.Marca='Almacen Global'" & _
IIf(line = "", "", " AND Cotizaciones.Linea='" & line & "'") & _
" AND Cotizaciones.Conversion=" & conv & " AND Cotizaciones.CodIngrediente='" & inge & "')"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PorcionGlobalDePorcion = rst("CodCotizacion")
rst.Close

Exit Function

Nulo:
PorcionGlobalDePorcion = 0

End Function

Function FrutaDePorcion(inge As String) As Long
'buscar la fruta conv 0 de cotizacion porcion
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.Marca, Cotizaciones.CodCotizacion, Cotizaciones.CodIngrediente, Cotizaciones.Conversion" & _
" FROM Cotizaciones" & _
" WHERE (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((Cotizaciones.Conversion)=0))" & _
" ORDER BY Cotizaciones.Marca DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FrutaDePorcion = rst("CodCotizacion")
rst.Close

Exit Function

Nulo:
FrutaDePorcion = 0

End Function

Function PesoIngredientes(bat As String) As Double

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Double

miSQL = "SELECT SubReceta.CodBatido, SubReceta.Cantidad, DBIngredientes.Unidad, DBIngredientes.Tipo, SubReceta.Tipo" & _
" FROM SubReceta INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE ((SubReceta.CodBatido='" & bat & "') AND (DBIngredientes.Tipo='Base' Or DBIngredientes.Tipo='Frutas' Or" & _
" DBIngredientes.Tipo='Saborizantes' Or DBIngredientes.Tipo='Verduras') AND (SubReceta.Tipo='B'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst
PesoIngredientes = 0

For I = 1 To Cant
    
    Select Case rst("Unidad")
        Case "gr", "mL"
        PesoIngredientes = PesoIngredientes + rst("Cantidad")
        Case "hojas"
        PesoIngredientes = PesoIngredientes + rst("Cantidad") * 10
        Case "medx30ml"
        PesoIngredientes = PesoIngredientes + rst("Cantidad") * 30
        Case "ramas"
        PesoIngredientes = PesoIngredientes + rst("Cantidad") * 10
        Case "sobre"
        PesoIngredientes = PesoIngredientes + rst("Cantidad") * 30
        Case "Unid"
        PesoIngredientes = PesoIngredientes + rst("Cantidad") * 80
        Case Else
        PesoIngredientes = PesoIngredientes + rst("Cantidad")
    End Select
    rst.MoveNext
Next I

rst.Close
Exit Function

AgainAgain:
PesoIngredientes = 0
End Function

Function ConcursoActivo() As Boolean

ConcursoActivo = DLookup("[ConcursoActivo]", "DatosSistema")

End Function

Function ProgramaEstaAbierto(process As String) As Boolean
    Dim objList As Object

    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        ProgramaEstaAbierto = True
    Else
        ProgramaEstaAbierto = False
    End If

End Function

Function AccesoDirectoBoletaVenta()
If Not CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
    DoCmd.OpenForm "Nota de Pedido", acNormal, , , acFormAdd
    [Forms]![Nota de Pedido]![Texto166] = 0
    'Call FotoClientePedido([Forms]![Nota de Pedido]![CodPedido])
    DoCmd.OpenForm "Menu PITAYA Global"
    Else
    MsgBox "Ya existe una boleta de venta abierta"
End If
End Function

Sub configurarmargentamañoreporte(repo As String)
    
Dim prt As Printer
Dim rpt2 As Access.Report

DoCmd.OpenReport repo, acDesign
Set prt = Reports(repo).Report.Printer
prt.TopMargin = 0
prt.BottomMargin = 0
prt.LeftMargin = 0
prt.RightMargin = 0
prt.DefaultSize = True 'tamaño de columna, columna es 1 asi que el tamaño tiene que ser por defecto mismo que papel
'MsgBox prt.PaperSize

'guardar cambios

'Make a random change
Set rpt2 = Reports(repo).Report
With rpt2
 If .DefaultView = 0 Then
     .DefaultView = 1
 Else
     .DefaultView = 0
 End If
    'Set it back
 If .DefaultView = 0 Then
     .DefaultView = 1
 Else
     .DefaultView = 0
 End If
End With
Set rpt2 = Nothing

'Set Reports(repo).Report.Printer = prt
DoCmd.Close acReport, repo, acSaveYes
End Sub

Sub configurarmargentamañoformulario(formu As String)
    
Dim prt As Printer
Dim rpt2 As Access.Form

DoCmd.OpenForm formu
Set prt = Forms(formu).Form.Printer
prt.TopMargin = 0
prt.BottomMargin = 0
prt.LeftMargin = 0
prt.RightMargin = 0
prt.DefaultSize = True 'tamaño de columna, columna es 1 asi que el tamaño tiene que ser por defecto mismo que papel


'guardar cambios

'Make a random change
'Set rpt2 = Forms(formu).Form
'With rpt2
' If .DefaultView = 0 Then
'     .DefaultView = 1
' Else
'     .DefaultView = 0
' End If
'    'Set it back
' If .DefaultView = 0 Then
'     .DefaultView = 1
' Else
'     .DefaultView = 0
' End If
'End With
'Set rpt2 = Nothing

'Set Reports(repo).Report.Printer = prt
'DoCmd.Close acReport, formu, acSaveYes
End Sub

Function ultimocodigopreingreso() As Long
'ultimo codigode preingreso
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodPreIngresoPitaya FROM PreIngresoPitaya ORDER BY CodPreIngresoPitaya DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimocodigopreingreso = rst("CodPreIngresoPitaya")
rst.Close

Exit Function

Nulo:
ultimocodigopreingreso = 0

End Function

Function mensajeaccesosdirectos() As String
MsgBox "F1: Lista de Productos" & vbCrLf & _
       "F2: Nuevo Pedido" & vbTab & vbCrLf & _
       "F3: Lista de Accesos Directos" & vbCrLf & _
       "F4: Marcar Entrada / Salida" & vbCrLf & _
       "F5: Actualizar" & vbCrLf & _
       "F6: Lista de Proovedores" & vbTab & vbCrLf & _
       "F7: Registro de Compras" & vbCrLf & _
       "F8: Registro de Ingresos" & vbCrLf & _
       "F9: Promociones Vigentes" & vbCrLf & _
       "F10: Minimizar Ventana Inicial" & vbCrLf & _
       "F11: Planilla" & vbCrLf & _
       "F12: Imprimir Etiqueta Porciones"
End Function
Function proyecciondiamesventasacumulado(di As Date, acum As Long) As Double
'acumulad dia a dia de un mes para llegar a acumulado total al final delmes
On Error GoTo Nulo

Dim diasa As Integer
diasa = Day(DateSerial(Year(di), Month(di) + 1, 0))
proyecciondiamesventasacumulado = (acum / diasa) * Day(di)
Exit Function

Nulo:
proyecciondiamesventasacumulado = 0

End Function

Function BuscarCodSubporcionamiento(codi As Long) As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPorcionamiento.CodProcesamiento, SubPorcionamiento.CodSubPorcionamiento" & _
" FROM SubPorcionamiento WHERE (((SubPorcionamiento.CodProcesamiento)=" & codi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
BuscarCodSubporcionamiento = rst("CodSubPorcionamiento")
Exit Function

Nulo:
BuscarCodSubporcionamiento = 0
End Function

Function UltimoCodSubporcionamiento() As Long
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPorcionamiento.CodSubPorcionamiento FROM SubPorcionamiento ORDER BY SubPorcionamiento.CodSubPorcionamiento DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
UltimoCodSubporcionamiento = rst("CodSubPorcionamiento")
Exit Function

Nulo:
UltimoCodSubporcionamiento = 0
End Function

Function absolutopositivo(dat As Double)
If dat < 0 Then
    absolutopositivo = 0
Else
    absolutopositivo = dat
End If
End Function

Function absolutoabsoluto(dat As Double)
If dat < 0 Then
    absolutoabsoluto = dat * (-1)
Else
    absolutoabsoluto = dat
End If
End Function

Function sucursalestotales() As Long
'TOTAL DE SUCURSALES GUARDADOS EN TABLA MIXED
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.CodSistema" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" ORDER BY DatosSistema.CodSistema DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
sucursalestotales = rst("CodSistema")
rst.Close

Exit Function

Nulo:
sucursalestotales = 1

End Function

Function primersaludo() As String
If Hour(Now) < 12 Then
    primersaludo = "Buenos Dias "
ElseIf Hour(Now) < 18 Then
    primersaludo = "Buenas Tardes "
Else
    primersaludo = "Buenas Noches "
End If
End Function

Function nombrediaespanol(fech As Date) As String
'nombre del dia en espanol
On Error GoTo Nulo
Select Case Weekday(fech, 2)
Case 1
    nombrediaespanol = "Lunes"
Case 2
    nombrediaespanol = "Martes"
Case 3
    nombrediaespanol = "Miercoles"
Case 4
    nombrediaespanol = "Jueves"
Case 5
    nombrediaespanol = "Viernes"
Case 6
    nombrediaespanol = "Sabado"
Case 7
    nombrediaespanol = "Domingo"
Case Else
    nombrediaespanol = ""
End Select

Exit Function
Nulo:
nombrediaespanol = ""
End Function

Function ultimoresumen() As Long
'ultimo codigo de resumen
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT codresumenpago FROM resumenpago ORDER BY codresumenpago DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimoresumen = rst("codresumenpago")
rst.Close

Exit Function

Nulo:
ultimoresumen = 0

End Function
Function ultimaordendecompra() As Long
'ultimo codigo de orden de compra
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT codordendecompra FROM OrdenDeCompra ORDER BY codordendecompra DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
ultimaordendecompra = rst("codordendecompra")
rst.Close

Exit Function

Nulo:
ultimaordendecompra = 0

End Function

Function nombreciudadsucursalglobal(ciu As Integer) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Buscar Codigo CLUB
miSQL = "SELECT DatosSistema.Ciudad, DatosSistema.CodSistema FROM DatosSistema" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (DatosSistema.CodSistema=" & ciu & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreciudadsucursalglobal = rst("Ciudad")
rst.Close

Exit Function

Nulo:
nombreciudadsucursalglobal = "Sin Ciudad"

End Function


Function nombreciudadsucursalglobalPitayaX(pitayax As String) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Buscar Codigo CLUB
miSQL = "SELECT DatosSistema.Ciudad, DatosSistema.Nombre FROM DatosSistema" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (DatosSistema.Nombre='" & pitayax & "')"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
nombreciudadsucursalglobalPitayaX = rst("Ciudad")
rst.Close

Exit Function

Nulo:
nombreciudadsucursalglobalPitayaX = "Sin Ciudad"

End Function

Function SiglasInicialesTexto(atexto As String, siglas As Integer) As String

' Obtener las siglas del proveedor (primeras siglas letras de cada palabra)
Dim siglaspartido As String
Dim partido() As String
partido = Split(atexto, " ")
siglaspartido = ""

For Each palabra In partido
    siglaspartido = siglaspartido & Left(palabra, siglas)
Next palabra


SiglasInicialesTexto = siglaspartido

End Function

Function leyendacontroles() As String
leyendacontroles = "CM: Consumo Maximo Segun Historial - SI: Stock Inicial/Inventario Inicial - SF: Stock Final/Inventario FInal"
End Function

Function leyendabalances() As String
leyendabalances = "SI: Inventario Inicial - SF: Inventario FInal - IN: Ingresos - ME: Mermas - CI: Consumo x Inventario - CR: COnsumo x Recetas - V: Variacion"
End Function
Function leyendaproduccion() As String
leyendaproduccion = "SI: Inventario Inicial - SFT: Inventario FInal Teorico - SFR: Inventario FInal Real - TRAN: Resultado de Transformacion - CO: Compras - DE: Despacho"
End Function

Function leyendavariacion() As String
leyendavariacion = "V: Variacion (CI-CR)/CR  -->  Positivo: Faltante - Negativo: Sobrante"
End Function

Function leyendacompras() As String
leyendacompras = "CM: Consumo Maximo Segun Historial - SI: Stock Inicial/Inventario Inicial - PE: Pedido Total - CO: Compras"
End Function

Function tamanovasobatido(medi As String) As String
Select Case medi
    Case "Gigantona"
        tamanovasobatido = "20oz"
    Case "Mediano"
        tamanovasobatido = "16oz"
    Case "Kid"
        tamanovasobatido = "12oz"
    Case Else
        tamanovasobatido = ""
End Select
End Function
Function inicialesvasobatido(medi As String) As String
Select Case medi
    Case "Gigantona"
        inicialesvasobatido = "N"
    Case "Mediano"
        inicialesvasobatido = "P"
    Case "Kid"
        inicialesvasobatido = "K"
    Case Else
        inicialesvasobatido = ""
End Select
End Function
Function cotiinsumoclavexnombreproductoventa(nomb As String) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Buscar Codigo CLUB
miSQL = "SELECT codigoyversionproductoventa([DBBatidos]![CodBatido]) AS codigoversion," & _
" DBBatidos.Nombre, SubReceta.InsumoClave," & _
" IIf(IsNull([SubReceta]![codporcion]),cotidirectodeingrediente([Subreceta]![CodIngrediente]),[SubReceta]![codporcion]) AS cond" & _
" FROM (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente" & _
" WHERE (((DBBatidos.Nombre) = '" & nomb & "') And ((SubReceta.InsumoClave) = True))" & _
" ORDER BY codigoyversionproductoventa([DBBatidos]![CodBatido]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
cotiinsumoclavexnombreproductoventa = rst("cond")
rst.Close

Exit Function

Nulo:
cotiinsumoclavexnombreproductoventa = 0

End Function

Function codigoyversionproductoventa(caden As String) As Long
'On Error GoTo Nulo
Dim I As Integer
Dim Caracter As String
Dim grupocodigo As Long
Dim grupoversion As Long
Dim tempCaracter As String
Dim codigoprincipal As Integer
codigoprincipal = 0
grupocodigo = 0
grupoversion = 0
tempCaracter = ""
codigoyversionproductoventa = 0

For I = 1 To Len(caden)
    If IsNumeric(Mid(caden, I, 1)) Then
        Caracter = Mid(caden, I, 1)
        tempCaracter = tempCaracter & Caracter
        codigoprincipal = 1
    
    Else
        If codigoprincipal = 1 Then 'tenia numero y ahora encontro tecto entonces inmediatamente pasa a x 1000
            If grupocodigo = 0 Then
                grupocodigo = CInt(tempCaracter) * 100
            ElseIf grupoversion = 0 Then
                grupoversion = CInt(tempCaracter)
            End If
            tempCaracter = ""
            codigoprincipal = 0
        End If
    End If

Next

If tempCaracter <> "" Then 'hay codigo de version
    If grupocodigo = 0 Then
        grupocodigo = CInt(tempCaracter) * 100
    ElseIf grupoversion = 0 Then
        grupoversion = CInt(tempCaracter)
    End If
End If
codigoyversionproductoventa = grupocodigo + grupoversion

Exit Function

Nulo:

End Function
Function paroimpar(nume As Integer)
If nume Mod 2 = 0 Then
    paroimpar = "par"
Else
    paroimpar = "impar"
End If
End Function

Function elegirimpresora(Tipo As String) As String
On erro GoTo Nulo
' enlistar impresoras
For Each prtLoop In Application.Printers
    With prtLoop
        If .DeviceName Like "*" & DLookup("[" & Tipo & "]", "[SistemaGlobal]") & "*" Then
            elegirimpresora = .DeviceName
        End If
        '.DeviceName & ";" _
        '.DriverName & ";" _
        '.Port
    End With
Next prtLoop
Exit Function
Nulo:
MsgBox "impresora buscada no se encuentra instalada"
elegirimpresora = ""
End Function
Function corregirtextomalescrito(tex As String) As String
Dim textocorregido As String
textocorregido = tex

Revaluar:
If Left(textocorregido, 1) = " " Then
    textocorregido = Right(textocorregido, Len(textocorregido) - 1)
    GoTo Revaluar
End If

If Right(textocorregido, 1) = " " Then
    textocorregido = Left(textocorregido, Len(textocorregido) - 1)
    GoTo Revaluar
End If

Dim conta As Integer
Dim recorrido As String
Dim division As String
Dim totalloop As Integer
Dim arraytex() As String

totalloop = Len(textocorregido)
For conta = 1 To totalloop
    recorrido = Right(Left(textocorregido, conta), 1)
    If IsNumeric(recorrido) Then
        textocorregido = Left(textocorregido, conta - 1) & Right(textocorregido, Len(textocorregido) - conta)
        conta = conta - 1
    End If
Next conta

textocorregido = UCase(Left(textocorregido, 1)) & LCase(Right(textocorregido, Len(textocorregido) - 1))

corregirtextomalescrito = textocorregido

End Function

Function alturaconsultasql(consul As String, alto As Integer) As Long
'1 inch = 1440 twips, resultado en twips
'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
MsgBox consul
miSQL = consul
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
alturaconsultasql = rst.RecordCount * alto * 1440
MsgBox rst.RecordCount
rst.Close

Exit Function

Nulo:
alturaconsultasql = 0

End Function


Sub NotificacionConSonido(Tipo As String)

' Mensaje de notificación
'MsgBox "¡Tienes una nueva notificación!"

' Reproduce el archivo WAV
Dim rutaSonido As String
rutaSonido = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Sonidos\" & Tipo & ".wav"
Call PlaySound(rutaSonido, 0, SND_ASYNC Or SND_FILENAME)
End Sub

Function cantidadsucursalesexistentes() As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Buscar Codigo CLUB
miSQL = "SELECT StatusSucursales.CodLocal FROM StatusSucursales ORDER BY StatusSucursales.CodLocal DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
cantidadsucursalesexistentes = rst("CodLocal")
rst.Close

Exit Function

Nulo:
cantidadsucursalesexistentes = 1

End Function

Function BuscarRutaArchivo(nombreArchivoBuscado As String, carpetaRaiz As String) As String
Dim fso As Object
Dim carpeta As Object
Dim archivo As Object
Dim subcarpeta As Object
Dim rutaEncontrada As String

Set fso = CreateObject("Scripting.FileSystemObject")
Set carpeta = fso.GetFolder(carpetaRaiz)

' Buscar en archivos de la carpeta actual
For Each archivo In carpeta.Files
    If StrComp(archivo.Name, nombreArchivoBuscado, vbTextCompare) = 0 Then
        BuscarRutaArchivo = archivo.path
        Exit Function
    End If
Next archivo

' Buscar en subcarpetas
For Each subcarpeta In carpeta.SubFolders
    rutaEncontrada = BuscarRutaArchivo(nombreArchivoBuscado, subcarpeta.path)
    If rutaEncontrada <> "" Then
        BuscarRutaArchivo = rutaEncontrada
        Exit Function
    End If
Next subcarpeta

' Si no lo encuentra, devuelve cadena vacía
BuscarRutaArchivo = ""

End Function

Function cfechasqlfechahora(fechaaccess As Date) As String
cfechasqlfechahora = Format(fechaaccess, "yyyy-mm-dd hh:nn:ss")
End Function
Function cfechasqlfecha(fechaaccess As Date) As String
cfechasqlfecha = Format(fechaaccess, "yyyy-mm-dd")
End Function
Function cfechasqlhora(fechaaccess As Date) As String
cfechasqlhora = Format(fechaaccess, "hh:nn:ss")
End Function

Function esmoduloopitayaraiz() As Integer
    Select Case CurrentProject.Name
        Case "Pitaya_System.accdb" 'archivo base de sistema
            esmoduloopitayaraiz = 1
        Case "Pitaya_System(1).accdb" 'archivo base de sistema
            esmoduloopitayaraiz = 1
        Case "Pitaya_System - Copy.accdb" 'archivo base de sistema
            esmoduloopitayaraiz = 1
        Case "Pitaya_System - Copia.accdb" 'archivo base de sistema
            esmoduloopitayaraiz = 1
        Case Else ' otro archivo
            If CurrentProject.Name Like "Modulo*" Then
                esmoduloopitayaraiz = 1
            Else
                esmoduloopitayaraiz = 0
            End If
    End Select
    
End Function
