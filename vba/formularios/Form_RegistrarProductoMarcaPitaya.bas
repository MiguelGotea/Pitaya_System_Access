' ==========================================================
' Modulo  : Form_RegistrarProductoMarcaPitaya
' Tipo    : 100
' Lineas  : 451
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database

Private Sub Comando137_Click()
On Error GoTo Salir
'Pajilla  Doble
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)


DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (575, " & 2 * Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (437, " & 3 * Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando142_Click()
On Error GoTo Salir
'Pajilla  Dorada
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (575, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (437, " & 3 * Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando148_Click()
On Error GoTo Salir
'Pajilla  con cepillo
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (575, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (437, " & 3 * Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando301_Click()
On Error GoTo Salir
'Miel 260gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (796, " & Cant & ", #" & Me.fechaprocedencia & "#)" 'envase vidrio

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (797, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (798, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 797)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (798, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando342_Click()
On Error GoTo Salir
'COCOA EN POLVO
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (700, " & Cant & ", #" & Me.fechaprocedencia & "#)" ' empaque ziploc

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (819, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (818, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 819)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (818, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando357_Click()
On Error GoTo Salir
'Marañon 56gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (901, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (900, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 901)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (900, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando366_Click()
On Error GoTo Salir
'FRUTOS SDESHIDRATADOS 56GR
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (899, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (898, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 899)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (898, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando73_Click()
On Error GoTo Salir
'Galleta avena
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (820, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (821, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 820)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (821, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub


Private Sub Form_Close()
If CurrentProject.AllForms("Ingresos a Pitaya").IsLoaded Then
    [Forms]![Ingresos a Pitaya].Form.Requery
End If
If CurrentProject.AllForms("Ingreso Inventario Pitaya").IsLoaded Then
    [Forms]![Ingreso Inventario Pitaya].Form.Subformulario_Inventario_Cotizacion.Requery
End If
If CurrentProject.AllForms("RegistroPreIngresosPitaya").IsLoaded Then
    [Forms]![RegistroPreingresosPitaya].Form.Requery
End If
DoCmd.SetWarnings True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
DoCmd.SetWarnings False
End Sub

'''''''''''''''''''''''''''''''''''''''''NUEVA PRESENTACION
Private Sub Comando262_Click()
On Error GoTo Salir
'Granola 230gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (700, " & Cant & ", #" & Me.fechaprocedencia & "#)" 'empaque ziploc

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (747, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (711, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 747)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (711, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando177_Click()
On Error GoTo Salir
'Harina 200gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (700, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (737, " & 0.05 * Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (703, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 737)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (703, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando211_Click()
On Error GoTo Salir
'Cacao 151gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (700, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (822, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (701, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 822)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (701, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub



Private Sub Comando179_Click()
On Error GoTo Salir
'Pecanas 51gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (823, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (560, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 823)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (560, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando181_Click()
On Error GoTo Salir
'Marañon 76gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (824, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (705, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 824)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (705, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando202_Click()
On Error GoTo Salir
'Almendras 83gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (735, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (706, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 735)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (706, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando276_Click()
On Error GoTo Salir
'Semillas Mixtas 95gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    'Almendra
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (825, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (759, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 825)"
    'Mani
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (826, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (760, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 826)"
    'Pasas
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (827, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (761, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 827)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (759, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (760, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (761, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If


Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando180_Click()
On Error GoTo Salir
'Pistachos 65gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (436, " & Cant & ", #" & Me.fechaprocedencia & "#)"
DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (438, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (828, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (707, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 828)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (707, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub

Private Sub Comando271_Click()
On Error GoTo Salir
'Miel 550gr
Dim Cant As Integer
Cant = InputBox("Cantidad de Productos", "Ingresar Datos", 1)

DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
" (756, " & Cant & ", #" & Me.fechaprocedencia & "#)"

If Me.adesde = "[IngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (755, " & Cant & ", #" & Me.fechaprocedencia & "#)"
    DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
    " (762, 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", 755)"
Else
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Cantidad, Fecha) values" & _
    " (762, " & Cant & ", #" & Me.fechaprocedencia & "#)"
End If

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente, Intentelo nuevamente"
End Sub
