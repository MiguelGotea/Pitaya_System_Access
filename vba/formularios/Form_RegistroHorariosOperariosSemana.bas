' ==========================================================
' Modulo  : Form_RegistroHorariosOperariosSemana
' Tipo    : 100
' Lineas  : 391
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database
Public Sub modohorariobloqueado()
Me.CodTipoLunes.Locked = True
Me.TE1Lunes.Locked = True
Me.TS1Lunes.Locked = True

Me.CodTipoMartes.Locked = True
Me.TE1Martes.Locked = True
Me.TS1Martes.Locked = True

Me.CodTipoMiercoles.Locked = True
Me.TE1Miercoles.Locked = True
Me.TS1Miercoles.Locked = True

Me.CodTipoJueves.Locked = True
Me.TE1Jueves.Locked = True
Me.TS1Jueves.Locked = True

Me.CodTipoViernes.Locked = True
Me.TE1Viernes.Locked = True
Me.TS1Viernes.Locked = True

Me.CodTipoSabado.Locked = True
Me.TE1Sabado.Locked = True
Me.TS1Sabado.Locked = True

Me.CodTipoDomingo.Locked = True
Me.TE1Domingo.Locked = True
Me.TS1Domingo.Locked = True

Me.HorasFaltantes.Locked = True

Me.Comando910.Visible = False
Me.Comando949.Visible = False
End Sub

Public Sub modohorarionobloqueado()
Me.CodTipoLunes.Locked = False
Me.TE1Lunes.Locked = False
Me.TS1Lunes.Locked = False

Me.CodTipoMartes.Locked = False
Me.TE1Martes.Locked = False
Me.TS1Martes.Locked = False

Me.CodTipoMiercoles.Locked = False
Me.TE1Miercoles.Locked = False
Me.TS1Miercoles.Locked = False

Me.CodTipoJueves.Locked = False
Me.TE1Jueves.Locked = False
Me.TS1Jueves.Locked = False

Me.CodTipoViernes.Locked = False
Me.TE1Viernes.Locked = False
Me.TS1Viernes.Locked = False

Me.CodTipoSabado.Locked = False
Me.TE1Sabado.Locked = False
Me.TS1Sabado.Locked = False

Me.CodTipoDomingo.Locked = False
Me.TE1Domingo.Locked = False
Me.TS1Domingo.Locked = False

Me.HorasFaltantes.Locked = False

Me.Comando910.Visible = True
Me.Comando949.Visible = True
End Sub


Private Sub CodTipoLunes_Exit(Cancel As Integer)
If Me.CodTipoLunes <> 3 Then
    Me.TE1Lunes = ""
    Me.TS1Lunes = ""
End If
End Sub
Private Sub CodTipoMartes_Exit(Cancel As Integer)
If Me.CodTipoMartes <> 3 Then
    Me.TE1Martes = ""
    Me.TS1Martes = ""
End If
End Sub
Private Sub CodTipoMiercoles_Exit(Cancel As Integer)
If Me.CodTipoMiercoles <> 3 Then
    Me.TE1Miercoles = ""
    Me.TS1Miercoles = ""
End If
End Sub
Private Sub CodTipoJueves_Exit(Cancel As Integer)
If Me.CodTipoJueves <> 3 Then
    Me.TE1Jueves = ""
    Me.TS1Jueves = ""
End If
End Sub
Private Sub CodTipoViernes_Exit(Cancel As Integer)
If Me.CodTipoViernes <> 3 Then
    Me.TE1Viernes = ""
    Me.TS1Viernes = ""
End If
End Sub
Private Sub CodTipoSabado_Exit(Cancel As Integer)
If Me.CodTipoSabado <> 3 Then
    Me.TE1Sabado = ""
    Me.TS1Sabado = ""
End If
End Sub
Private Sub CodTipoDomingo_Exit(Cancel As Integer)
If Me.CodTipoDomingo <> 3 Then
    Me.TE1Domingo = ""
    Me.TS1Domingo = ""
End If
End Sub

Private Sub Comando155_Click()
Me.Requery
If Me.asemana <= numerosemana(Date) Then
    Call modohorariobloqueado
ElseIf Me.asemana = numerosemana(Date) + 1 And Weekday(Date, vbMonday) < 6 Then
    Call modohorarionobloqueado
ElseIf Me.asemana = numerosemana(Date) + 1 And Weekday(Date, vbMonday) >= 6 Then
    Call modohorariobloqueado
ElseIf Me.asemana > numerosemana(Date) + 1 Then
    Call modohorariobloqueado
Else
    Call modohorariobloqueado
End If
End Sub

Private Sub Comando731_Click()
DoCmd.OpenForm "Horario_Semana", acFormPivotTable, , "[Sucursal]=" & codigoLocal() & " AND [sema]=" & Me.asemana & " AND [CodOperario]=" & Me.CodOperario
[Forms]![Horario_Semana].Form.InsideWidth = 20000
[Forms]![Horario_Semana].Form.InsideHeight = 6800
End Sub
Private Sub Comando737_Click()
DoCmd.OpenForm "Horario_Semana", acFormPivotTable, , "[Sucursal]=" & codigoLocal() & " AND [sema]=" & Me.asemana & " AND [dianumero]=" & 2
[Forms]![Horario_Semana].Form.InsideWidth = 20000
[Forms]![Horario_Semana].Form.InsideHeight = 6800
End Sub


Private Sub Comando702_Click()
If Me.asemana <= numerosemana(Date) Then
    MsgBox "No se pueden editar horarios de semana actual o pasadas"
    Exit Sub
ElseIf Me.asemana = numerosemana(Date) + 1 And Weekday(Date, vbMonday) >= 6 Then
    MsgBox "No se pueden editar horarios, fecha limite viernes"
    Exit Sub
ElseIf Me.asemana > numerosemana(Date) + 1 Then
    MsgBox "No se peuden editar hoarios tan lejanos"
    Exit Sub
End If

If MsgBox("Generar la lista de colaboradores activos para asignarle un horario", vbYesNo, "Generar Horario") = vbNo Then
    Exit Sub
Else
    If IsNull(DLookup("[CodOperario]", "[HorariosSucursales]", "[semana]=" & Me.asemana)) Then
        'no existe horarios creso esta semana
    Else
        'si existe hroario
        MsgBox "Ya existe un horario creado para esta semana"
        Exit Sub
    End If
End If
Dim rst As DAO.Recordset
Dim miSQL As String
Dim semax As Integer
semax = Me.asemana

miSQL = "SELECT NombreOperario([Operarios]![CodOperario]) AS Name, Operarios.CodOperario, Operarios.Operativo," & _
" NivelesCargos.Marcacion, AsignacionNivelesCargos.Sucursal," & _
" numerosemana([AsignacionNivelesCargos]![Fecha])<=" & semax & " And IIf(IsNull([AsignacionNivelesCargos]![Fin]),Date(),[AsignacionNivelesCargos]![Fin])>=Date() AS condi" & _
" FROM (Operarios INNER JOIN AsignacionNivelesCargos ON Operarios.CodOperario = AsignacionNivelesCargos.CodOperario)" & _
" INNER JOIN NivelesCargos ON AsignacionNivelesCargos.CodNivelesCargos = NivelesCargos.CodNivelesCargos" & _
" GROUP BY NombreOperario([Operarios]![CodOperario]), Operarios.CodOperario, Operarios.Operativo, NivelesCargos.Marcacion," & _
" AsignacionNivelesCargos.Sucursal, numerosemana([AsignacionNivelesCargos]![Fecha])<=" & semax & "" & _
" And IIf(IsNull([AsignacionNivelesCargos]![Fin]),Date(),[AsignacionNivelesCargos]![Fin])>=Date()" & _
" HAVING (((Operarios.Operativo)<>0) AND ((NivelesCargos.Marcacion)<>0)" & _
" AND ((AsignacionNivelesCargos.Sucursal)=" & codigoLocal() & ") AND ((numerosemana([AsignacionNivelesCargos]![Fecha])<=" & semax & "" & _
" And IIf(IsNull([AsignacionNivelesCargos]![Fin]),Date(),[AsignacionNivelesCargos]![Fin])>=Date())<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO HorariosSucursales(CodOperario, semana, CodTipoLunes, CodTipoMartes, CodTipoMiercoles, CodTipoJueves, CodTipoViernes, CodTipoSabado, CodTipoDomingo)" & _
    " values (" & rst("CodOperario") & ", " & Me.asemana & ",3,3,3,3,3,3,3)"
    DoCmd.SetWarnings True
    rst.MoveNext
Loop
rst.Close
'condi: numerosemana([AsignacionNivelesCargos]![Fecha])<=[Formularios]![EleccionColaboradorVigente]![asemana] Y SiInm(EsNulo([AsignacionNivelesCargos]![Fin]),Fecha(),[AsignacionNivelesCargos]![Fin])>=Fecha()
Me.Requery
End Sub



Private Sub Comando910_Click()
DoCmd.OpenForm "EleccionColaboradorVigente"
[Forms]![EleccionColaboradorVigente]![atabla] = "HorarioSucursales"
[Forms]![EleccionColaboradorVigente]![asucursal] = codigoLocal()
[Forms]![EleccionColaboradorVigente]![asemana] = [Forms]![RegistroHorariosOperariosSemana]![asemana]
[Forms]![EleccionColaboradorVigente]![CodigoBusqueda] = [Forms]![RegistroHorariosOperariosSemana]![CodOperario]
End Sub

Private Sub Comando929_Click()

DoCmd.OpenForm "LogueoAutorizacion"
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(21, 18, Date)
[Forms]![LogueoAutorizacion]![CodigoBusqueda].Enabled = False
[Forms]![LogueoAutorizacion]![origenlogueo] = "DesbloquearHorarioSucursal"

End Sub

Private Sub Comando949_Click()
If MsgBox("Desea eliminar el turno del colaborador?", vbYesNo, "CONFIRMACION") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM HorariosSucursales WHERE CodHorariosSucursales =" & Me.CodHorariosSucursales
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
Call modohorariobloqueado
End Sub





Private Sub HorasFaltantes_Exit(Cancel As Integer)
If IsNull(Me.HorasFaltantes) Then
    Me.HorasFaltantes = 0
End If
End Sub

'LUNES
Private Sub TE1Lunes_Click()
If Me.CodTipoLunes <> 3 Then
    Me.CodTipoLunes.SetFocus
End If
End Sub
Private Sub TE2Lunes_Click()
If Me.CodTipoLunes <> 3 Then
    Me.CodTipoLunes.SetFocus
End If
End Sub
Private Sub TS1Lunes_Click()
If Me.CodTipoLunes <> 3 Then
    Me.CodTipoLunes.SetFocus
End If
End Sub
Private Sub TS2Lunes_Click()
If Me.CodTipoLunes <> 3 Then
    Me.CodTipoLunes.SetFocus
End If
End Sub

'MARTES
Private Sub TE1Martes_Click()
If Me.CodTipoMartes <> 3 Then
    Me.CodTipoMartes.SetFocus
End If
End Sub
Private Sub TS1Martes_Click()
If Me.CodTipoMartes <> 3 Then
    Me.CodTipoMartes.SetFocus
End If
End Sub
Private Sub TE2Martes_Click()
If Me.CodTipoMartes <> 3 Then
    Me.CodTipoMartes.SetFocus
End If
End Sub
Private Sub TS2Martes_Click()
If Me.CodTipoMartes <> 3 Then
    Me.CodTipoMartes.SetFocus
End If
End Sub

'MIERCOLES
Private Sub TE1Miercoles_Click()
If Me.CodTipoMiercoles <> 3 Then
    Me.CodTipoMiercoles.SetFocus
End If
End Sub
Private Sub TS1Miercoles_Click()
If Me.CodTipoMiercoles <> 3 Then
    Me.CodTipoMiercoles.SetFocus
End If
End Sub
Private Sub TE2Miercoles_Click()
If Me.CodTipoMiercoles <> 3 Then
    Me.CodTipoMiercoles.SetFocus
End If
End Sub
Private Sub TS2Miercoles_Click()
If Me.CodTipoMiercoles <> 3 Then
    Me.CodTipoMiercoles.SetFocus
End If
End Sub

'JUEVES
Private Sub TE1jueves_Click()
If Me.CodTipoJueves <> 3 Then
    Me.CodTipoJueves.SetFocus
End If
End Sub
Private Sub TS1jueves_Click()
If Me.CodTipoJueves <> 3 Then
    Me.CodTipoJueves.SetFocus
End If
End Sub
Private Sub TE2jueves_Click()
If Me.CodTipoJueves <> 3 Then
    Me.CodTipoJueves.SetFocus
End If
End Sub
Private Sub TS2jueves_Click()
If Me.CodTipoJueves <> 3 Then
    Me.CodTipoJueves.SetFocus
End If
End Sub

'VIERNES
Private Sub TE1Viernes_Click()
If Me.CodTipoViernes <> 3 Then
    Me.CodTipoViernes.SetFocus
End If
End Sub
Private Sub TS1Viernes_Click()
If Me.CodTipoViernes <> 3 Then
    Me.CodTipoViernes.SetFocus
End If
End Sub
Private Sub TE2Viernes_Click()
If Me.CodTipoViernes <> 3 Then
    Me.CodTipoViernes.SetFocus
End If
End Sub
Private Sub TS2Viernes_Click()
If Me.CodTipoViernes <> 3 Then
    Me.CodTipoViernes.SetFocus
End If
End Sub

'SABADO
Private Sub TE1Sabado_Click()
If Me.CodTipoSabado <> 3 Then
    Me.CodTipoSabado.SetFocus
End If
End Sub
Private Sub TS1Sabado_Click()
If Me.CodTipoSabado <> 3 Then
    Me.CodTipoSabado.SetFocus
End If
End Sub
Private Sub TE2Sabado_Click()
If Me.CodTipoSabado <> 3 Then
    Me.CodTipoSabado.SetFocus
End If
End Sub
Private Sub TS2Sabado_Click()
If Me.CodTipoSabado <> 3 Then
    Me.CodTipoSabado.SetFocus
End If
End Sub

'DOMINGO
Private Sub TE1Domingo_Click()
If Me.CodTipoDomingo <> 3 Then
    Me.CodTipoDomingo.SetFocus
End If
End Sub
Private Sub TS1Domingo_Click()
If Me.CodTipoDomingo <> 3 Then
    Me.CodTipoDomingo.SetFocus
End If
End Sub
Private Sub TE2Domingo_Click()
If Me.CodTipoDomingo <> 3 Then
    Me.CodTipoDomingo.SetFocus
End If
End Sub
Private Sub TS2Domingo_Click()
If Me.CodTipoDomingo <> 3 Then
    Me.CodTipoDomingo.SetFocus
End If
End Sub
