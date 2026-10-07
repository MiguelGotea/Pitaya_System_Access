' ==========================================================
' Modulo  : Form_OpcionesCambioIngrediente
' Tipo    : 100  |  Lineas: 96
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Private Sub Comando403_Click()
Dim aped As Long
Dim obse As String
Dim codsubpe As Long
Dim nuevaobse As String
Dim codadicional As String
Dim precioadicional As Double
Dim conver As Double
Dim cantporcionesadicionalreceta As Double
Dim cantporcionesadicionalcambio As Double
Dim nombreporcionesadicionalreceta As String
Dim nombreporcionesadicionalcambio As String
Dim adicionaltieneporcioncambio As Integer

'Codigo de batido de adicional que cambiara a ingrediente de receta
codadicional = AdicionalVigenteDeIngrediente(Me.CodIngrediente)

' precio de adicional que cambiara a ingrediente de receta
precioadicional = DLookup("[Precio]", "[DBBatidos]", "[CodBatido]='" & codadicional & "'")


'codigo de pedido de nota de pedido
aped = [Forms]![CambiarIngredientesFactura]![apedido]

'codigo subpoedido de subpedido
codsubpe = [Forms]![CambiarIngredientesFactura]![asubpedido]

'observaciones donde se anexara texto
obse = DLookup("[Observaciones]", "[SubPedido]", "[CodSubPedido]=" & codsubpe)

'cantidad de porciones dentro de del adicional de cambio, del insumo con el que se cambiara
nombreporcionesadicionalcambio = nombreinsumoreceta(DLookup("[CodSubReceta]", "[SubReceta]", "[CodBatido]='" & Me.CodBatido & "' AND [InsumoClave]<>0"))

'cantidad de porciones dentro de del adicional de procedencia, del insumo que se cambiara
nombreporcionesadicionalreceta = nombreinsumoreceta(DLookup("[CodSubReceta]", "[SubReceta]", "[CodBatido]='" & Me.acodadicional & "' AND [InsumoClave]<>0"))

'cantidad de porciones dentro de del adicional de cambio, del insumo con el que se cambiara
cantporcionesadicionalcambio = cantidadinsumorecetaantesconvertirfraccion(DLookup("[CodSubReceta]", "[SubReceta]", "[CodBatido]='" & Me.CodBatido & "' AND [InsumoClave]<>0"))

'cantidad de porciones dentro de del adicional de procedencia, del insumo que se cambiara
cantporcionesadicionalreceta = cantidadinsumorecetaantesconvertirfraccion(DLookup("[CodSubReceta]", "[SubReceta]", "[CodBatido]='" & Me.acodadicional & "' AND [InsumoClave]<>0"))

'condicional si el adiconal es porcion o no
adicionaltieneporcioncambio = Nz(DLookup("[codporcion]", "[SubReceta]", "[CodBatido]='" & Me.CodBatido & "' AND [InsumoClave]<>0"), 0)

'euivalencia del valor total de porciones que sumen la cantiad de insumo a cambiar / valor de procion de insumo por cambiar
'cantidad de porciones de insumo nuevo equivalente al total de porciones de insumo de receta original segun precio total de sus porciones adicionles
If precioadicional = 0 Or IsNull(precioadicional) Then
    conver = 1
Else
    conver = Me.Texto720 / precioadicional
    'texto720=[aprecioadicional]*[equivalente](Precio de adicional de insumo de receta a cambiar* contenido de adicional / contenido de receta de insumo a cambiar
    
    'aplica si es porcion porque no se pueden agregar fraccion de porciones
    If adicionaltieneporcioncambio <> 0 Then
        conver = redondear_mas(conver * cantporcionesadicionalcambio) / cantporcionesadicionalcambio
        'si un adicional de cambio tiene 2 porc, multiplifca conver*2 lo redondea y lo vuelve a dividir en 2 , sale decimal pero en prociones seran completos
    End If
End If

'cantidadproductos= productos completos facturados batidos, bowl etc
DoCmd.SetWarnings False

''' Se agregan y quitan adicionales equivalentes a nivel de prociones para restar misma cantidad de prociones y sumar cantidad de porciones redonda

'Porcion nueva d ecambio
DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, SinAzucar, Azucar, CodPromocion, Vinculo)" & _
" VALUES ('" & Me.CodBatido & "', " & 1 * conver & ", " & aped & ", -1 , ' ', 92, " & codsubpe & ")"

'Porcion antigua que se elimina
DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, SinAzucar, Azucar, CodPromocion, Vinculo)" & _
" VALUES ('" & Me.acodadicional & "', " & -1 * Me.equivalente & ", " & aped & ", -1, ' ', 92, " & codsubpe & ")"

'Agregar a observaciones
nuevaobse = obse & ", AGREGAR " & fraccion(conver * cantporcionesadicionalcambio) & " " & nombreporcionesadicionalcambio & _
", QUITAR " & fraccion(Me.equivalente * cantporcionesadicionalreceta) & " " & nombreporcionesadicionalreceta

DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.Observaciones = '" & nuevaobse & "'" & _
" WHERE SubPedido.CodSubPedido = " & codsubpe

DoCmd.SetWarnings True

MsgBox "Cambio Realizado Exitosamente"

DoCmd.Close
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

