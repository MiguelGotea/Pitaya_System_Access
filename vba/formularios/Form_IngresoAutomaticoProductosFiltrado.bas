' ==========================================================
' Modulo  : Form_IngresoAutomaticoProductosFiltrado
' Tipo    : 100  |  Lineas: 137
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.aproducto.Text = "" Then
    If Me.adesdeform = "[Porcionamiento de Insumos]" And codigoLocal() <> 0 Then
        Me.Filter = "[NombreSinProcesar] like '*" & Me.aproducto.Text & "*' And [SubProducto]=0 And [TIPO1]='VARIABLES' And [Vigente]<>0 And [Descontinuado]=0 And [Marca2]<>'Almacen Global '"
    Else
        Me.Filter = "[NombreSinProcesar] like '*" & Me.aproducto.Text & "*'"
    End If
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
On Error GoTo Nulo

Dim cantix As Double
cantix = InputBox("Ingresar la cantidad: ", "Cantidad de productos", 0)

If cantix = 0 Then
    MsgBox "NO se agrego ningun producto"
    Exit Sub
End If

Select Case Me.adestino

Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
    " values (" & Me.CodCotizacion & ", " & cantix & ", " & Me.apreingreso & ")"
    DoCmd.SetWarnings True
    
Case "[IngresosPitaya]", "[Inventario Cotizacion]", "[Merma Cotizacion]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, Cantidad, Fecha)" & _
    " values (" & Me.CodCotizacion & ", " & cantix & ", #" & Me.afechapro & "#)"
    DoCmd.SetWarnings True
    
Case "[SubPorcionamiento]" 'centana de porcionamiento sellados
        
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(Procedencia, CodProcesamiento, Cantidad, Fecha)" & _
    " values (" & Me.CodCotizacion & ", 0, " & cantix & ", #" & Me.afechapro & "#)"
    DoCmd.SetWarnings True
    
Case "[Procesamiento]" ' agregar productos de procesamiento

    If Me.Conversion <> 0 Then
        MsgBox "NO se pueden agregar productos que no sean procesables"
        Exit Sub
    End If

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO " & Me.adestino & "(CodCotizacion, Cantidad, MedidaFinal, Fecha, Operario)" & _
    " values (" & Me.CodCotizacion & ", " & cantix & ", 0, #" & Me.afechapro & "#, 0)"
    DoCmd.SetWarnings True
    
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


