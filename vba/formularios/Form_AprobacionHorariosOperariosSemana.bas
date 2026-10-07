' ==========================================================
' Modulo  : Form_AprobacionHorariosOperariosSemana
' Tipo    : 100
' Lineas  : 328
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database
Private Sub asucursal_Exit(Cancel As Integer)
Me.acolaborador.Requery
End Sub
Private Sub acolaborador_Exit(Cancel As Integer)
If IsNull(Me.acolaborador) Then
    Me.ahistorial = 1
    Me.ahistorial.Enabled = False
Else
    Me.ahistorial.Enabled = True
End If
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
On Error Resume Next
Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "HorariosSucursales", "HorariosSucursales", 1)
Me.Requery

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
Private Sub Comando939_Click()
If MsgBox("Desea publicar horario oficial, una vez publicado ya no se podran hacer cambios", vbYesNo, "CONFIRMACION") = vbYes Then
    'se sube el hroario
Else
    Exit Sub
End If
Me.Requery

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantix As Integer

miSQL = "SELECT HorariosSucursales.semana, HorariosSucursales.CodOperario," & _
" HorariosSucursales.CodTipoLunes, HorariosSucursales.TE1Lunes, HorariosSucursales.TS1Lunes, HorariosSucursales.NotaLunes," & _
" HorariosSucursales.CodTipoMartes, HorariosSucursales.TE1Martes, HorariosSucursales.TS1Martes, HorariosSucursales.NotaMartes," & _
" HorariosSucursales.CodTipoMiercoles, HorariosSucursales.TE1Miercoles,  HorariosSucursales.TS1Miercoles, HorariosSucursales.NotaMiercoles," & _
" HorariosSucursales.CodTipoJueves, HorariosSucursales.TE1Jueves, HorariosSucursales.TS1Jueves, HorariosSucursales.NotaJueves," & _
" HorariosSucursales.CodTipoViernes, HorariosSucursales.TE1Viernes, HorariosSucursales.TS1Viernes, HorariosSucursales.NotaViernes," & _
" HorariosSucursales.CodTipoSabado, HorariosSucursales.TE1Sabado, HorariosSucursales.TS1Sabado, HorariosSucursales.NotaSabado," & _
" HorariosSucursales.CodTipoDomingo, HorariosSucursales.TE1Domingo, HorariosSucursales.TS1Domingo, HorariosSucursales.NotaDomingo," & _
" HorariosSucursales.HorasFaltantes" & _
" FROM HorariosSucursales WHERE (((HorariosSucursales.semana)=" & Me.asemana & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantix = rst.RecordCount
rst.MoveFirst

For I = 1 To cantix
    If I >= Me.icoti Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO HorariosSucursalesAprobados(CodOperario," & _
        " NotaLunes, CodTipoLunes, " & IIf(IsNull(rst("TE1Lunes")), "", "TE1Lunes, ") & IIf(IsNull(rst("TS1Lunes")), "", "TS1Lunes, ") & _
        " NotaMartes, CodTipoMartes, " & IIf(IsNull(rst("TE1Martes")), "", "TE1Martes, ") & IIf(IsNull(rst("TS1Martes")), "", "TS1Martes, ") & _
        " NotaMiercoles, CodTipoMiercoles, " & IIf(IsNull(rst("TE1Miercoles")), "", "TE1Miercoles, ") & IIf(IsNull(rst("TS1Miercoles")), "", "TS1Miercoles, ") & _
        " NotaJueves, CodTipoJueves, " & IIf(IsNull(rst("TE1Jueves")), "", "TE1Jueves, ") & IIf(IsNull(rst("TS1Jueves")), "", " TS1Jueves, ") & _
        " NotaViernes, CodTipoViernes, " & IIf(IsNull(rst("TE1Viernes")), "", "TE1Viernes, ") & IIf(IsNull(rst("TS1Viernes")), "", "TS1Viernes, ") & _
        " NotaSabado, CodTipoSabado, " & IIf(IsNull(rst("TE1Sabado")), "", "TE1Sabado, ") & IIf(IsNull(rst("TS1Sabado")), "", "TS1Sabado, ") & _
        " NotaDomingo, CodTipoDomingo, " & IIf(IsNull(rst("TE1Domingo")), "", "TE1Domingo, ") & IIf(IsNull(rst("TS1Domingo")), "", "TS1Domingo, ") & _
        " semana, HorasFaltantes, Sucursal, Aprobado)" & _
        " values (" & rst("CodOperario") & "," & _
        " '" & IIf(IsNull(rst("NotaLunes")), "", rst("NotaLunes")) & "', " & rst("CodTipoLunes") & ", " & IIf(IsNull(rst("TE1Lunes")), "", "#" & rst("TE1Lunes") & "#, ") & IIf(IsNull(rst("TS1Lunes")), "", "#" & rst("TS1Lunes") & "#, ") & _
        " '" & IIf(IsNull(rst("NotaMartes")), "", rst("NotaMartes")) & "', " & rst("CodTipoMartes") & ", " & IIf(IsNull(rst("TE1Martes")), "", "#" & rst("TE1Martes") & "#, ") & IIf(IsNull(rst("TS1Martes")), "", "#" & rst("TS1Martes") & "#, ") & _
        " '" & IIf(IsNull(rst("NotaMiercoles")), "", rst("NotaMiercoles")) & "', " & rst("CodTipoMiercoles") & ", " & IIf(IsNull(rst("TE1Miercoles")), "", "#" & rst("TE1Miercoles") & "#, ") & IIf(IsNull(rst("TS1Miercoles")), "", "#" & rst("TS1Miercoles") & "#, ") & _
        " '" & IIf(IsNull(rst("NotaJueves")), "", rst("NotaJueves")) & "', " & rst("CodTipoJueves") & ", " & IIf(IsNull(rst("TE1Jueves")), "", "#" & rst("TE1Jueves") & "#, ") & IIf(IsNull(rst("TS1Jueves")), "", "#" & rst("TS1Jueves") & "#, ") & _
        " '" & IIf(IsNull(rst("NotaViernes")), "", rst("NotaViernes")) & "', " & rst("CodTipoViernes") & ", " & IIf(IsNull(rst("TE1Viernes")), "", "#" & rst("TE1Viernes") & "#, ") & IIf(IsNull(rst("TS1Viernes")), "", "#" & rst("TS1Viernes") & "#, ") & _
        " '" & IIf(IsNull(rst("NotaSabado")), "", rst("NotaSabado")) & "', " & rst("CodTipoSabado") & ", " & IIf(IsNull(rst("TE1Sabado")), "", "#" & rst("TE1Sabado") & "#, ") & IIf(IsNull(rst("TS1Sabado")), "", "#" & rst("TS1Sabado") & "#, ") & _
        " '" & IIf(IsNull(rst("NotaDomingo")), "", rst("NotaDomingo")) & "', " & rst("CodTipoDomingo") & ", " & IIf(IsNull(rst("TE1Domingo")), "", "#" & rst("TE1Domingo") & "#, ") & IIf(IsNull(rst("TS1Domingo")), "", "#" & rst("TS1Domingo") & "#, ") & _
        " " & Me.asemana & ", " & IIf(IsNull(rst("HorasFaltantes")), 0, rst("HorasFaltantes")) & ", " & Me.asucursal & ", #" & Now & "#)"
        DoCmd.SetWarnings True
    End If
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
MsgBox "Horario revisado y publicado de " & nombrelocalglobal(Me.asucursal)
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call Comando939_Click
End Sub

Private Sub Comando949_Click()
If MsgBox("Desea eliminar el turno del colaborador?", vbYesNo, "CONFIRMACION") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM HorariosSucursales WHERE CodHorariosSucursales =" & Me.CodHorariosSucursales
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Comando973_Click()
DoCmd.OpenForm "EleccionColaboradorVigente"
[Forms]![EleccionColaboradorVigente]![atabla] = "HorarioSucursales"
[Forms]![EleccionColaboradorVigente]![asucursal] = [Forms]![AprobacionHorariosOperariosSemana]![asucursal]
[Forms]![EleccionColaboradorVigente]![asemana] = [Forms]![AprobacionHorariosOperariosSemana]![asemana]
[Forms]![EleccionColaboradorVigente]![CodigoBusqueda] = [Forms]![AprobacionHorariosOperariosSemana]![CodOperario]
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
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
