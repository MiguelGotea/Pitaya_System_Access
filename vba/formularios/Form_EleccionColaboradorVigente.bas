' ==========================================================
' Modulo  : Form_EleccionColaboradorVigente
' Tipo    : 100  |  Lineas: 36
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database






Private Sub Comando48_Click()
If MsgBox("Desea ingresar colaborador al horario?", vbYesNo, "CONFIRMACION") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO HorariosSucursales(CodOperario, semana, CodTipoLunes, CodTipoMartes, CodTipoMiercoles, CodTipoJueves, CodTipoViernes, CodTipoSabado, CodTipoDomingo)" & _
    " values (" & Me.CodigoBusqueda & ", " & Me.asemana & ",3,3,3,3,3,3,3)"
    DoCmd.SetWarnings True
    
    If CurrentProject.AllForms("AprobacionHorariosOperariosSemana").IsLoaded Then
        [Forms]![AprobacionHorariosOperariosSemana].Form.Requery
    End If
    
    If CurrentProject.AllForms("RegistroHorariosOperariosSemana").IsLoaded Then
        [Forms]![RegistroHorariosOperariosSemana].Form.Requery
    End If
    DoCmd.Close
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Call importartablasweb


End Sub



