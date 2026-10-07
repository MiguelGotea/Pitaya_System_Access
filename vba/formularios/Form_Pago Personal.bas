' ==========================================================
' Modulo  : Form_Pago Personal
' Tipo    : 100
' Lineas  : 348
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database

Private Sub Comando198_Click()
If IsNull(Me.CodigoBusqueda) = True Then
    Me.Filter = "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "#"
    Me.FilterOn = True
Else
    Me.Filter = "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "# and [CodOperario]= " & Me.CodigoBusqueda
    Me.FilterOn = True
End If
End Sub

Public Sub Comando262_Click()
If codigoLocal() = 0 Then 'Sistema local
    Exit Sub
End If

Me.Filter = "semana = " & Me.asemana & " and minutos < 15 and minutos > 0"
Me.FilterOn = True

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.asemana & " - 9.ReporteDeTardanzas.pdf"
DoCmd.OutputTo acOutputForm, "Pago Personal", acFormatPDF, archivo, False

End Sub

Private Sub Comando403_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM RegistroHorario WHERE CodHorario =" & Me.CodHorario
DoCmd.SetWarnings True
Me.Requery
End Sub

Private Sub Comando420_Click()
If IsNull(Me.CodigoBusqueda) = True Then
    MsgBox ("Ingresar Operario")
Else
    Dim anoi, mesi As Long
    anoi = Year(Date - 200)
    mesi = Month(Date - 200)
    DoCmd.OpenForm "Busqueda Compras", , , "[Mes]+[Año]*100>=" & mesi + anoi * 100 & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Personal' AND [Observaciones] like '*" & DLookup("[Nombre]", "[Operarios]", "[CodOperario]=" & Me.CodigoBusqueda) & "*'"
End If
End Sub

Private Sub Comando422_Click()
DoCmd.OpenForm "Horario_Semana", acFormPivotTable, , "[CodOperario]=" & Me.CodigoBusqueda & " AND [Dates]<=#" & Me.hasta & "# AND [Dates]>=#" & Me.desde & "#"
[Forms]![Horario_Semana].Form.InsideWidth = 17000
[Forms]![Horario_Semana].Form.InsideHeight = 7500
End Sub

Private Sub Comando438_Click()
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Double

miSQL = "SELECT Operarios.CodOperario, Operarios.Operativo, Operarios.Sucursal FROM Operarios" & _
" WHERE (((Operarios.Operativo)<>0) AND ((Operarios.Sucursal)=codigolocal()) AND ((Operarios.Cargo)='Lider' OR (Operarios.Cargo)='Operario'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For I = 1 To canti
    Me.CodigoBusqueda = rst("CodOperario")
    Call Comando48_Click
    rst.MoveNext
Next I

rst.Close
Exit Sub

Nulo:
MsgBox "Hay algun dato erroneo, verificar"
End Sub

Private Sub Comando441_Click()

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Double
Dim numtel As String
Dim mensa As String

Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta & "\" & Me.mesboleta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

miSQL = "SELECT Operarios.CodOperario, Operarios.Operativo, Operarios.Sucursal FROM Operarios" & _
" WHERE (((Operarios.Operativo)<>0) AND ((Operarios.Sucursal)=codigolocal()) AND ((Operarios.Cargo)='Lider' OR (Operarios.Cargo)='Operario'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For I = 1 To canti
    Me.CodigoBusqueda = rst("CodOperario")
    numtel = DLookup("Celular", "Operarios", "[CodOperario]=" & Me.CodigoBusqueda)
    mensa = primersaludo() & NombreOperario(Me.CodigoBusqueda) & ", revisar los turnos marcados esta quincena"
    'archivo = Direccionmes & "\" & Me.intervaloboleta & " - " & codigolocal() & " - " & NombreOperario(Me.CodigoBusqueda) & " " & ".jpg"
    archivo = Direccionmes & "\temporal.png"
    DoCmd.OpenForm "Resumen Horas Personal", acFormPivotTable, , "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "# and [CodOperario]= " & Me.CodigoBusqueda
    [Forms]![Resumen Horas Personal].Form.Requery
    [Forms]![Resumen Horas Personal].PivotTable.exportpicture archivo, , 400, 600
    DoCmd.Close acForm, "Resumen Horas Personal"
    Call EnviarImagenWhatsapp(mensa, archivo, numtel)
    Kill (archivo)
    rst.MoveNext
    
Next I

rst.Close
Exit Sub

Nulo:
MsgBox "Hay algun dato erroneo, verificar"
End Sub



Private Sub Comando48_Click()
Call rellenarturnosvaciosoperarios(Me.CodigoBusqueda, Me.desde, Me.hasta)
Me.Requery
Dim Direccionano, Direccionmes As String
Dim archivo As String
Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta & "\" & Me.mesboleta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If
archivo = Direccionmes & "\" & Me.intervaloboleta & " - " & codigoLocal() & " - " & NombreOperario(Me.CodigoBusqueda) & " " & ".pdf"
DoCmd.OpenReport "Reporte Pago Personal", acViewReport, , "[CodOperario]=" & Me.CodigoBusqueda
[Reports]![Reporte Pago Personal].OrderBy = "[Fecha]"
[Reports]![Reporte Pago Personal].OrderByOn = True
[Reports]![Reporte Pago Personal].Report.Requery
DoCmd.OutputTo acOutputReport, "Reporte Pago Personal", acFormatPDF, archivo, False
DoCmd.Close acReport, "Reporte Pago Personal"
End Sub

Private Sub Comando497_Click()
On Error GoTo Nulo
Dim canti As Double
Dim numtel As String
Dim mensa As String

Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta & "\" & Me.mesboleta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

numtel = DLookup("Celular", "Operarios", "[CodOperario]=" & Me.CodigoBusqueda)
mensa = primersaludo() & NombreOperario(Me.CodigoBusqueda) & " revisar los turnos marcados esta quincena"
'archivo = Direccionmes & "\" & Me.intervaloboleta & " - " & codigolocal() & " - " & NombreOperario(Me.CodigoBusqueda) & " " & ".jpg"
archivo = Direccionmes & "\temporal.png"
DoCmd.OpenForm "Resumen Horas Personal", acFormPivotTable, , "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "# and [CodOperario]= " & Me.CodigoBusqueda
[Forms]![Resumen Horas Personal].Form.Requery
[Forms]![Resumen Horas Personal].PivotTable.exportpicture archivo, , 400, 600
DoCmd.Close acForm, "Resumen Horas Personal"

Call EnviarImagenWhatsapp(mensa, archivo, numtel)
Kill (archivo)
Exit Sub

Nulo:
MsgBox "Hay algun dato erroneo, verificar"
End Sub

Private Sub Comando500_Click()
DoCmd.OpenForm "RegistroHorarios"
End Sub

Private Sub Comando502_Click()
DoCmd.OpenForm "RegistroVacaciones"
End Sub



Private Sub Comando526_Click()
DoCmd.OpenForm "Horario_Semana", acFormPivotTable, , "[Operativo]<>0 AND [Sucursal]= " & codigoLocal() & " AND [Dates]<=#" & Me.hasta & "# AND [Dates]>=#" & Me.desde & "#"
[Forms]![Horario_Semana].Form.InsideWidth = 17000
[Forms]![Horario_Semana].Form.InsideHeight = 7500
End Sub

Private Sub Comando529_Click()
DoCmd.OpenForm "CumpleanosEquipoPitaya"
End Sub





Private Sub Comando579_Click()
DoCmd.OpenForm "RegistroCompensacionFeriados"
End Sub

Private Sub Comando584_Click()

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Double
Dim numtel As String
Dim mensa As String

Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta & "\" & Me.mesboleta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

miSQL = "SELECT Operarios.CodOperario, Operarios.Operativo, Operarios.Sucursal FROM Operarios" & _
" WHERE (((Operarios.Operativo)<>0) AND ((Operarios.Sucursal)=codigolocal()) AND ((Operarios.Cargo)='Lider' OR (Operarios.Cargo)='Operario'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For I = 1 To canti
    Me.CodigoBusqueda = rst("CodOperario")
    numtel = DLookup("Celular", "Operarios", "[CodOperario]=" & Me.CodigoBusqueda)
    mensa = primersaludo() & NombreOperario(Me.CodigoBusqueda) & ", revisar los turnos marcados esta quincena"
    'archivo = Direccionmes & "\" & Me.intervaloboleta & " - " & codigolocal() & " - " & NombreOperario(Me.CodigoBusqueda) & " " & ".jpg"
    archivo = Direccionmes & "\" & Me.intervaloboleta & " - " & codigoLocal() & " - " & NombreOperario(Me.CodigoBusqueda) & ".png"
    DoCmd.OpenForm "Resumen Horas Personal", acFormPivotTable, , "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "# and [CodOperario]= " & Me.CodigoBusqueda
    [Forms]![Resumen Horas Personal].Form.Requery
    [Forms]![Resumen Horas Personal].PivotTable.exportpicture archivo, , 400, 600
    DoCmd.Close acForm, "Resumen Horas Personal"
    rst.MoveNext
    
Next I

rst.Close
Exit Sub

Nulo:
MsgBox "Hay algun dato erroneo, verificar"
End Sub

Private Sub Comando586_Click()
On Error GoTo Error

Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Me.añoboleta & "\" & Me.mesboleta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

Call Comando198_Click

DoCmd.OutputTo acOutputForm, "Pago Personal", acFormatXLSX, Direccionmes & "\" & Me.intervaloboleta & " - " & codigoLocal() & " - " & NombreOperario(Me.CodigoBusqueda) & ".xlsx", False
MsgBox "Datos Exportados"

Exit Sub
Error:
MsgBox "Ingresar Datos Correctamente, Revisar si el excel esta actualmente abierto"
End Sub

Private Sub Comando710_Click()
DoCmd.OpenForm "RegistroSubsidioPersonal"
End Sub

Private Sub Comando729_Click()
DoCmd.OpenForm "PlanillaPersonal"
End Sub

Private Sub Comando746_Click()
On Error GoTo Error

Dim Direccionano As String
Dim Direccionmes As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Year(Me.desde)
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo RRHH\Boletas Operarios\" & Year(Me.desde) & "\" & Month(Me.desde)
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

DoCmd.OutputTo acOutputQuery, "HorasCumplidasSemanalSucursalExcel", acFormatXLSX, Direccionmes & "\" & codigoLocal() & " - " & Day(Me.desde) & " al " & Day(Me.hasta) & " - Horas totales por semana total quincena.xlsx", False
MsgBox "Datos Exportados en la carpeta del mes " & Month(Me.desde)
Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Call importartablasweb
End Sub

Private Sub intervalo_Change()
Me.desde.Requery
Me.hasta.Requery

End Sub


