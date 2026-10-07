' ==========================================================
' Modulo  : Form_CambiarIngredientesFactura
' Tipo    : 100  |  Lineas: 119
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Private Sub Comando403_Click()
Dim codadi As String
Dim conver As Double
Dim cantidaddentroadicional As Double
Dim cantidadporcambio As Double

'Codigo de adicional en DBBatidos
codadi = AdicionalVigenteDeIngrediente(Me.CodIngrediente)

'Abrir ventana de opcines de cambio
DoCmd.OpenForm "OpcionesCambioIngrediente"

''Enviar a ventana de opciones OpcionesCambioIngrediente

'DBIngredientes.Nombre nombre de producto
[Forms]![OpcionesCambioIngrediente]![anombre] = Me.Nombre

'SubReceta.Cantidad : cantidad de producto a cambiar
cantidadporcambio = Me.Cantidad
[Forms]![OpcionesCambioIngrediente]![acontenido] = cantidadporcambio

'Detalle de cantidad de insumo en la receta que se cambiara
'[Forms]![OpcionesCambioIngrediente]![cantidadrecetaparacambio] = Me.cantireceta & " " & Me.unidadreceta & " " & Me.nombrereceta

'Enviar cantidad de productos facturados enteros, batidos completos, bowl wtc
[Forms]![OpcionesCambioIngrediente]![cantidadproductos] = Me.cantidadproductos

'Codigo de batido de adicional seleccionado
[Forms]![OpcionesCambioIngrediente]![acodadicional] = codadi

'Contenido de ingrediente especifico del adiional seleciconado
cantidaddentroadicional = DLookup("[Cantidad]", "[SubReceta]", "[CodBatido]='" & codadi & "' AND [CodIngrediente]='" & Me.CodIngrediente & "'")
[Forms]![OpcionesCambioIngrediente]![acontenidoadicional] = cantidaddentroadicional

'precio de adicional seleccionado
[Forms]![OpcionesCambioIngrediente]![aprecioadicional] = DLookup("[Precio]", "[DBBatidos]", "[CodBatido]='" & codadi & "'") 'precio del adicioanl seleccionado

'cantidad de proucto de receta a cambiar / contenido de ingrediente en adicional = cantidad porciones adicioanles que equivalen a contenido de receta por cambiar
If cantidaddentroadicional = 0 Or IsNull(cantidaddentroadicional) Then
    conver = 0
Else
    conver = cantidadporcambio / cantidaddentroadicional
End If
[Forms]![OpcionesCambioIngrediente]![equivalente] = conver

[Forms]![OpcionesCambioIngrediente].Form.Requery
End Sub

Private Sub Comando657_Click()
Dim codadicional As String
Dim obse As String
Dim codsubpe As Long
Dim nuevaobse As String
Dim equivalente As Double
Dim cantidaddentroadicional As Double
Dim nombreporcionesadicionalreceta As String

'subpedido
codsubpe = Me.asubpedido

'Observacion existente
obse = DLookup("[Observaciones]", "[SubPedido]", "[CodSubPedido]=" & codsubpe)

'COdigo de adicional relacionado a ingrediente
codadicional = AdicionalVigenteDeIngrediente(Me.CodIngrediente)

'cantidad de insumo ingrediente dentro del adicional encontrado de ese ingrediente
cantidaddentroadicional = DLookup("[Cantidad]", "[SubReceta]", "[CodBatido]='" & codadicional & "' AND [CodIngrediente]='" & Me.CodIngrediente & "'")

'SubReceta.Cantidad : cantidad de producto a cambiar
cantidadporcambio = Me.Cantidad

'cantidad de proucto de receta a cambiar / contenido de ingrediente en adicional = cantidad porciones adicioanles que equivalen a contenido de receta por cambiar
If cantidaddentroadicional = 0 Or IsNull(cantidaddentroadicional) Then
    equivalente = 0
Else
    equivalente = cantidadporcambio / cantidaddentroadicional
End If

'cantidad de porciones dentro de del adicional de procedencia, del insumo que se cambiara
nombreporcionesadicionalreceta = nombreinsumoreceta(DLookup("[CodSubReceta]", "[SubReceta]", "[CodBatido]='" & codadicional & "' AND [InsumoClave]<>0"))

'cantidad de porciones dentro de del adicional de procedencia, del insumo que se cambiara
cantporcionesadicionalreceta = cantidadinsumorecetaantesconvertirfraccion(DLookup("[CodSubReceta]", "[SubReceta]", "[CodBatido]='" & codadicional & "' AND [InsumoClave]<>0"))

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, SinAzucar, Azucar, CodPromocion, Vinculo)" & _
" VALUES ('" & codadicional & "', " & -1 * equivalente & ", " & Me.apedido & ", -1, ' ', 92, " & codsubpe & ")"

nuevaobse = obse & ", QUITAR " & fraccion(equivalente * cantporcionesadicionalreceta) & " " & nombreporcionesadicionalreceta

DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.Observaciones = '" & nuevaobse & "' WHERE SubPedido.CodSubPedido = " & codsubpe

DoCmd.SetWarnings True

MsgBox "Cambio Realizado Exitosamente"

End Sub

Private Sub Form_Close()


If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![Nota de Pedido].Form.Secundario57.Requery
    [Forms]![Nota de Pedido].Form.Texto1278.Requery

    [Forms]![Nota de Pedido].Form.recibecor.SetFocus
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

