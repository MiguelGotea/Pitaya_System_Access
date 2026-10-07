' ==========================================================
' Modulo  : Form_VerificarMarcacionesSemana
' Tipo    : 100  |  Lineas: 29
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database



  
Private Sub Comando148_Click()
If IsNull(Me.asucursal) Or Me.asucursal = "" Then
    MsgBox "Ingresar Sucursal"
    Exit Sub
End If
If IsNull(Me.asemana) Or Me.asemana = "" Then
    MsgBox "Ingresar Semana"
    Exit Sub
End If


Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "RegistroHorario", "RegistroHorario", 1)

Call importartablaespecifica("SupervisionSucursales", "HorariosSucursalesAprobados", "HorariosSucursalesAprobados", 1)
Me.Requery

End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub
