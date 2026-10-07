' ==========================================================
' Modulo  : Form_HistorialPreIngresosLocal
' Tipo    : 100  |  Lineas: 320
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database
Private Sub crearvalidacionpreingreso(codi As Long)
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantix As Integer

Dim ex As Long

miSQL = "SELECT SubPreIngresosPitaya.CodSubPreIngresoPitaya, SubPreIngresosPitaya.CodPreIngresoPitaya" & _
" FROM SubPreIngresosPitaya WHERE (((SubPreIngresosPitaya.CodPreIngresoPitaya)=" & codi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantix = rst.RecordCount
rst.MoveFirst

For ex = 1 To cantix
    If ex >= Me.icoti Then
        
        If IsNull(DLookup("[CodSubPreIngresoPitaya]", "[ValidacionPreIngreso]", "[CodSubPreIngresoPitaya]=" & rst("CodSubPreIngresoPitaya"))) Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO ValidacionPreIngreso(CodSubPreIngresoPitaya) values (" & rst("CodSubPreIngresoPitaya") & ")"
            DoCmd.SetWarnings True
        Else
            'Ya existe registro de validacion para esta linea de preingreso
        End If
        
    End If
    rst.MoveNext
Next ex

rst.Close

Me.icoti = 1
Exit Sub

AgainAgain:
Me.icoti = ex
rst.Close
Call crearvalidacionpreingreso(codi)
End Sub


Public Sub modosucursal()
Me.modali = "modosucursal"
Me.Comando196.Visible = False  'Boton de nuevo preingreso
Me.Comando277.Visible = False  'Boton de eliminar

Me.Etiqueta260.Visible = True  'etiqeuta stuatus
Me.Status.Visible = True 'valor status

Me.Etiqueta303.Visible = False ';visible etiqyetar fecha
Me.Fecha.Visible = False ';visible r fecha
Me.Fecha.width = 0.1 ';ancho r fecha
Me.Fecha.Locked = True ';no cambiar fecha

Me.Destino.Locked = True 'no cambiar destino

Me.Etiqueta321.Visible = False  'etiqeuta validado
Me.valida.Visible = False 'valor validado
Me.valida.width = 0.1  'valor ancho de validado

Me.Etiqueta333.Visible = False  'etiqeuta existen cambios de sucursal
Me.existencambios.Visible = False 'valor existen cambios de sucursal
Me.existencambios.width = 0.1 'ancho existen cambios de sucursal

Me.Form.Filter = "[Destino] = '" & nombrelocal() & "' AND [valida]='VALIDADO'"
Me.Form.FilterOn = True
Me.Form.Requery
End Sub

Public Sub modosucursaladministracion()
Me.modali = "modosucursaladministracion"
Me.Comando196.Visible = False  'Boton de nuevo preingreso
Me.Comando277.Visible = False  'Boton de eliminar

Me.Etiqueta260.Visible = True  'etiqeuta stuatus
Me.Status.Visible = True 'valor status

Me.Etiqueta303.Visible = False ';visible etiqyetar fecha
Me.Fecha.Visible = False ';visible r fecha
Me.Fecha.width = 0.1 ';ancho r fecha
Me.Fecha.Locked = True ';no cambiar fecha

Me.Destino.Locked = True 'no cambiar destino

Me.Etiqueta321.Visible = True  'etiqeuta validado
Me.valida.Visible = True 'valor validado
'Me.valida.Width = 1  'valor ancho de validado

Me.Etiqueta333.Visible = False  'etiqeuta existen cambios de sucursal
Me.existencambios.Visible = False 'valor existen cambios de sucursal
Me.existencambios.width = 0.1 'ancho existen cambios de sucursal

Me.Form.Filter = "[Destino] = '" & nombrelocal() & "'"
Me.Form.FilterOn = True
Me.Form.Requery
End Sub

Public Sub modocentral()
Me.modali = "modocentral"
Me.Comando196.Visible = False  'Boton de nuevo preingreso
Me.Comando277.Visible = False  'Boton de eliminar

Me.Etiqueta260.Visible = False  'etiqeuta stuatus
Me.Status.Visible = False 'valor status
Me.Status.width = 0.1  'valor ancho de stoatus

Me.Etiqueta303.Visible = False ';visible etiqyetar fecha
Me.Fecha.Visible = False ';visible r fecha
Me.Fecha.width = 0.1 ';ancho r fecha
Me.Fecha.Locked = True ';no cambiar fecha

Me.Etiqueta321.Visible = True  'etiqeuta validado
Me.valida.Visible = True 'valor validado

Me.Etiqueta333.Visible = False  'etiqeuta existen cambios de sucursal
Me.existencambios.Visible = False 'valor existen cambios de sucursal
Me.existencambios.width = 0.1 'ancho existen cambios de sucursal

Me.Destino.Locked = True 'no cambiar destino

'Todos los preingresos visibles
End Sub

Public Sub modopredespacho()
Me.modali = "modopredespacho"
Me.Comando196.Visible = False  'Boton de nuevo preingreso
Me.Comando277.Visible = False  'Boton de eliminar

Me.Etiqueta260.Visible = False  'etiqeuta stuatus
Me.Status.Visible = False 'valor status
Me.Status.width = 0.1  'valor ancho de stoatus

Me.Etiqueta303.Visible = False ';visible etiqyetar fecha
Me.Fecha.Visible = False ';visible r fecha
Me.Fecha.width = 0.1 ';ancho r fecha
Me.Fecha.Locked = True ';no cambiar fecha

Me.Etiqueta321.Visible = True  'etiqeuta validado
Me.valida.Visible = True 'valor validado

Me.Etiqueta333.Visible = False  'etiqeuta existen cambios de sucursal
Me.existencambios.Visible = False 'valor existen cambios de sucursal
Me.existencambios.width = 0.1 'ancho existen cambios de sucursal

Me.Destino.Locked = True 'no cambiar destino

'Todos los preingresos visibles
End Sub

Public Sub mododespacho()
Me.modali = "mododespacho"
Me.Comando196.Visible = True  'Boton de nuevo preingreso
Me.Comando277.Visible = True  'Boton de eliminar

Me.Etiqueta260.Visible = False  'etiqeuta stuatus
Me.Status.Visible = False 'valor status
Me.Status.width = 0.1  'valor ancho de stoatus

Me.Etiqueta303.Visible = True ';visible etiqyetar fecha
Me.Fecha.Visible = True ';visible r fecha
Me.Fecha.Locked = False ';no cambiar fecha\

Me.Etiqueta321.Visible = True  'etiqeuta validado
Me.valida.Visible = True 'valor validado

Me.Etiqueta333.Visible = True  'etiqeuta existen cambios de sucursal
Me.existencambios.Visible = True 'valor existen cambios de sucursal
'Me.existencambios.Width = 0.1 'ancho existen cambios de sucursal

Me.Destino.Locked = False 'no cambiar destino

'Todos los preingresos visibles
End Sub


Private Sub Comando112_Click()
fechaac.Value = fechaac.Value + 1
actualizarediciondatos
Me.Requery
End Sub

Sub actualizarediciondatos()
If Me.fechaac < Date Then
    Me.Fecha.Enabled = False
    Me.Destino.Enabled = False
Else
    Me.Fecha.Enabled = True
    Me.Destino.Enabled = True
End If

End Sub

Private Sub Comando113_Click()
fechaac.Value = fechaac.Value - 1
actualizarediciondatos
Me.Requery
End Sub

Private Sub Comando148_Click()
Dim desti As String
Dim locaf As Integer
Dim vali As Integer
desti = DLookup("[Destino]", "[PreIngresoPitaya]", "[CodPreingresoPitaya]=" & Me.CodPreIngresoPitaya)
locaf = CInt(Split(desti, " ")(1))
vali = DLookup("[Validado]", "[PreIngresoPitaya]", "[CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya)

Select Case Me.modali

Case "modopredespacho"
    Call crearvalidacionpreingreso(Me.CodPreIngresoPitaya)
    
    DoCmd.OpenForm "ValidacionPreIngresosPitaya", , , "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
                   " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
    [Forms]![ValidacionPreIngresosPitaya]![aencabezado].Caption = "VALIDACION DE PREINGRESO " & UCase(desti) & " " & UCase(nombreciudadsucursalglobal(locaf))
    [Forms]![ValidacionPreIngresosPitaya]![fechaac] = Me.Fecha
    [Forms]![ValidacionPreIngresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
    [Forms]![ValidacionPreIngresosPitaya]![codigoope] = Me.codigoope
    [Forms]![ValidacionPreIngresosPitaya]![validado] = IIf(vali = 0, "EN REVISION", "VALIDADO")

Case "mododespacho"
    
    DoCmd.OpenForm "RegistroPreIngresosPitaya", , , "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
                   " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
    [Forms]![RegistroPreingresosPitaya]![aencabezado].Caption = "REGISTRO DE PREINGRESO " & UCase(desti) & " " & UCase(nombreciudadsucursalglobal(locaf))
    [Forms]![RegistroPreingresosPitaya]![fechaac] = Me.Fecha
    [Forms]![RegistroPreingresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
    [Forms]![RegistroPreingresosPitaya]![codigoope] = Me.codigoope
    [Forms]![RegistroPreingresosPitaya]![validado] = IIf(vali = 0, "EN REVISION", "VALIDADO")
    
    If vali = 0 Then ' no validado puede entrar modo edicion
        Call Forms("[RegistroPreIngresosPitaya]").modovistadespacho
    Else
        Call Forms("[RegistroPreIngresosPitaya]").modovistacentral
    End If
    
Case "modocentral"
    
    DoCmd.OpenForm "RegistroPreIngresosPitaya", , , "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
                   " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
    [Forms]![RegistroPreingresosPitaya]![aencabezado].Caption = "DETALLE DE PREINGRESO " & UCase(desti) & " " & UCase(nombreciudadsucursalglobal(locaf))
    [Forms]![RegistroPreingresosPitaya]![fechaac] = Me.Fecha
    [Forms]![RegistroPreingresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
    [Forms]![RegistroPreingresosPitaya]![codigoope] = Me.codigoope
    [Forms]![RegistroPreingresosPitaya]![validado] = IIf(vali = 0, "EN REVISION", "VALIDADO")

    Call Forms("[RegistroPreIngresosPitaya]").modovistacentral
  
Case "modosucursal"
    
    DoCmd.OpenForm "RegistroPreIngresosPitaya", , , "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
                   " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
    [Forms]![RegistroPreingresosPitaya]![aencabezado].Caption = "REVISION DE PREINGRESO " & UCase(desti) & " " & UCase(nombreciudadsucursalglobal(locaf))
    [Forms]![RegistroPreingresosPitaya]![fechaac] = Me.Fecha
    [Forms]![RegistroPreingresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
    [Forms]![RegistroPreingresosPitaya]![codigoope] = Me.codigoope
    [Forms]![RegistroPreingresosPitaya]![validado] = IIf(vali = 0, "EN REVISION", "VALIDADO")
    [Forms]![RegistroPreingresosPitaya]![Status] = IIf(statuspreingreso(Me.CodPreIngresoPitaya) = 0, "PENDIENTE", "INGRESADO")

    Call Forms("[RegistroPreIngresosPitaya]").modovistasucursal


Case "modosucursaladministracion"
    
    DoCmd.OpenForm "RegistroPreIngresosPitaya", , , "[SubPreIngresosPitaya]![CodPreIngresoPitaya]=" & Me.CodPreIngresoPitaya & _
                   " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
    [Forms]![RegistroPreingresosPitaya]![aencabezado].Caption = "REVISION DE PREINGRESO " & UCase(desti) & " " & UCase(nombreciudadsucursalglobal(locaf))
    [Forms]![RegistroPreingresosPitaya]![fechaac] = Me.Fecha
    [Forms]![RegistroPreingresosPitaya]![acodigo] = Me.CodPreIngresoPitaya
    [Forms]![RegistroPreingresosPitaya]![codigoope] = Me.codigoope
    [Forms]![RegistroPreingresosPitaya]![validado] = IIf(vali = 0, "EN REVISION", "VALIDADO")
    [Forms]![RegistroPreingresosPitaya]![Status] = IIf(statuspreingreso(Me.CodPreIngresoPitaya) = 0, "PENDIENTE", "INGRESADO")

    Call Forms("[RegistroPreIngresosPitaya]").modovistasucursaladministracion

Case Else
    'No abre nada
    
End Select

End Sub



Private Sub Comando196_Click()
DoCmd.OpenForm "DefinirDatosPreingreso"
End Sub

Private Sub Comando277_Click()
If ContadorItemsPreingreso(Me.CodPreIngresoPitaya) = 0 Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM PreIngresoPitaya WHERE CodPreIngresoPitaya = " & Me.CodPreIngresoPitaya
    DoCmd.SetWarnings True
    Me.Requery
Else
    MsgBox "Eliminar items dentro del preingreso para poder eliminar el preingreso"
End If

End Sub

Private Sub Fecha_Exit(Cancel As Integer)
If Me.Fecha < Date Then
    MsgBox "No se puede mover despacho a fechas pasadas"
    Me.Fecha = Me.fechaac
End If
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub fechaac_Exit(Cancel As Integer)
Call actualizarediciondatos
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
End Sub
