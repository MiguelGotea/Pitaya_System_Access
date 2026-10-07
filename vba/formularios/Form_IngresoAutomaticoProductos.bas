' ==========================================================
' Modulo  : Form_IngresoAutomaticoProductos
' Tipo    : 100  |  Lineas: 148
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.aproducto.Text = "" Then
    Me.Filter = "[NombreSinProcesar] like '*" & Me.aproducto.Text & "*' And [Vigente]<>0 And [Descontinuado] =0 And [Marca2]<>'Almacen Global '" & IIf(Me.modulo = "contabilidad", " And [SubProducto]=0", "")
    Me.FilterOn = True
    Me.aproducto.SelStart = Len(Me.aproducto)
Else
    Me.FilterOn = False
    Me.aproducto.SetFocus
End If

Exit Sub

NoResultados:
MsgBox "No se encontraron coincidencias"
Me.FilterOn = False
Me.aproducto.SetFocus
Me.aproducto.Text = Left(Me.aproducto.Text, Len(Me.aproducto.Text) - 1)
Call actualizarlista

End Sub
'******************************** INICIO DE FORMULAS DE BOTONES Y CONTROLES ******************
'
'Borrar datos cuando se selecciona un tipo de buscador



Private Sub aproducto_GotFocus()

Me.imagen.Picture = ""
'Me.FilterOn = False
End Sub

'Buscador de datos
Private Sub aproducto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
    Me.aproducto.Text = Left(Me.aproducto.Text, Len(Me.aproducto.Text) - 1)
    Call actualizarlista
Else
    Call actualizarlista
End If
End Sub





Private Sub Comando309_Click()
'On Error GoTo Nulo

Dim Cant As Double

If Me.adestino <> "[Compras]" Then
    Cant = InputBox("Ingresar la cantidad: ", "Cantidad de productos", 0)
End If

Select Case Me.adestino

Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
    " values (" & Me.CodCotizacion & ", " & Cant & ", " & Me.apreingreso & ")"
    DoCmd.SetWarnings True
    
Case "[IngresosPitaya]", "[Inventario Cotizacion]", "Merma Cotizacion"

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, Cantidad, Fecha)" & _
    " values (" & Me.CodCotizacion & ", " & Cant & ", #" & Me.afechapro & "#)"
    DoCmd.SetWarnings True
    
Case "[SubOrdenDeCompra]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, cantidadorden, codordendecompra)" & _
    " values (" & Me.CodCotizacion & ", " & Cant & ", " & Me.apreingreso & ")"
    DoCmd.SetWarnings True
    
Case "[Compras]"
    Dim horaInsertar As Date
    
    ' Verificar si el formulario "Cierre por Turno 2" está abierto
    If CurrentProject.AllForms("Cierre por Turno 2").IsLoaded Then
        horaInsertar = [Forms]![Cierre por Turno 2]![aHoraFinal] - TimeSerial(0, 5, 0)
    Else
        horaInsertar = Time()
    End If
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, Cantidad, Fecha, CostoTotal, CodProveedor, Destino, Tipo, NumeroFactura, CodOperario, Hora)" & _
    " values (" & Me.CodCotizacion & ", 1, #" & Me.afechapro & "#, 0, " & Me.compraproovedor & ", '" & "Pitaya " & codigoLocal() & "', '" & Me.compratipopago & "', '" & Me.comprafactura & "', " & Me.compraoperario & ", #" & horaInsertar & "#)"
    DoCmd.SetWarnings True
    
    [Forms]![RegistroDeFacturaCompra].Form.Requery
    
    If [Forms]![RegistroDeFacturaCompra]![tipofacturaform] = "NUEVO" Then
    [Forms]![RegistroDeFacturaCompra]![Comando851].Enabled = True
    [Forms]![RegistroDeFacturaCompra]![aoperario].Enabled = False
    [Forms]![RegistroDeFacturaCompra]![aproovedor].Enabled = False
    [Forms]![RegistroDeFacturaCompra]![afactura].Enabled = False
    [Forms]![RegistroDeFacturaCompra]![atipopago].Enabled = False
    [Forms]![RegistroDeFacturaCompra]![Comando924].Enabled = False
    
    End If
    
Case Else
    
    MsgBox "No se encuentra ninguna ventana abierta donde ingresar datos"
    
End Select

Exit Sub
Nulo:
MsgBox "Ingresar los datos correctamente"
End Sub

Private Sub Form_Close()
If Me.adesdeform = "[Ingreso Inventario Almacen]" Then
    Forms(Me.adesdeform).Form.Subformulario_Inventario_Cotizacion.Requery
End If
Forms(Me.adesdeform).Form.Requery

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub nombrex_DblClick(Cancel As Integer)

DoCmd.OpenForm "Analisis de Productos", acNormal
[Forms]![Analisis de Productos]![CodigoBusqueda] = [Forms]![Relacion de Cotizaciones]![CodCotizacion]
End Sub

'*************************** FOTO DE PRODUCTO ****************
Private Sub nombrex_GotFocus()
If Not Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Productos\" & Me.CodCotizacion & ".png") = "" Then
    Me.imagen.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Productos\" & Me.CodCotizacion & ".png"
    Else
    Me.imagen.Picture = ""
End If
End Sub


