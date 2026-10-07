' ==========================================================
' Modulo  : Form_DefinirDatosPreingreso
' Tipo    : 100
' Lineas  : 57
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database


Private Sub botonbloquear_Click()
On Error GoTo Nulo
If IsNull(Me.fechaac) Then
    MsgBox "Ingresar fecha de despacho"
    Exit Sub
End If
If IsNull(Me.Sucursal) Then
    MsgBox "Ingresar sucursal"
    Exit Sub
End If
'Nuevo preingreso
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO PreIngresoPitaya(Fecha, Hora, Destino)" & _
" values (#" & Me.fechaac & "#, #" & Time & "#,'" & Me.Sucursal & "')"
DoCmd.SetWarnings True

Dim codpre As Long
codpre = ultimocodigopreingreso()
DoCmd.OpenForm "RegistroPreIngresosPitaya", , , "[CodPreIngresoPitaya]=" & codpre & _
    " AND PorcionDentroDeMezcla([SubPreIngresosPitaya]![CodCotizacion])=0"
               
[Forms]![RegistroPreingresosPitaya]![aencabezado].Caption = "REGISTRO DE PREINGRESO " & UCase(DLookup("[Nombre]", "[StatusSucursales]", "[CodLocal] =" & Right([Forms]![DefinirDatosPreingreso]![Sucursal], Len([Forms]![DefinirDatosPreingreso]![Sucursal]) - 7)))
[Forms]![RegistroPreingresosPitaya]![fechaac] = Me.fechaac
[Forms]![RegistroPreingresosPitaya]![acodigo] = codpre

DoCmd.Close acForm, "DefinirDatosPreingreso"
Exit Sub
Nulo:
MsgBox "No se ha creado registro, vuelva a intentarlo"

End Sub

Private Sub Comando1281_Click()
DoCmd.Close
End Sub

Private Sub fechaac_Exit(Cancel As Integer)
If Me.fechaac < Date Then
    MsgBox "No se pueden crear preingresos con fechas pasadas"
    Me.fechaac = Date
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False




End Sub



