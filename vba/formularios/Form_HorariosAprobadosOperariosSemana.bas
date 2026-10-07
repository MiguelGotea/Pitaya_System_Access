' ==========================================================
' Modulo  : Form_HorariosAprobadosOperariosSemana
' Tipo    : 100  |  Lineas: 38
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database





Private Sub Comando1015_Click()
If Me.asemana = "" Or IsNull(Me.asemana) Then
    MsgBox "Ingresar numero de semana del horairo que se eliminara"
    Exit Sub
End If
If Me.asucursal = "" Or IsNull(Me.asucursal) Then
    MsgBox "Ingresar sucursal del horairo que se eliminara"
    Exit Sub
End If

If MsgBox("Desea eliminar el horario aprobado y publicado? se tendra que subir nuevamente el horario", vbYesNo, "Eliminar Horario Aprobado") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM HorariosSucursalesAprobados WHERE Sucursal =" & Me.asucursal & " AND semana=" & Me.asemana
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Comando155_Click()

Me.Requery
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
End Sub




