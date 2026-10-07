' ==========================================================
' Modulo  : Form_IngresoDePorciones
' Tipo    : 100  |  Lineas: 37
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando14_Click()

On Error GoTo Salir

Dim codsupor As Long
Dim Cant As Integer
Cant = InputBox("Cantidad de Porciones", "Ingresar Datos", 1)
If Cant <= 0 Then
    MsgBox "Dato Erroneo"
    Exit Sub
End If

DoCmd.SetWarnings False
'Ingreso
DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values (" & Me.CodCotizacion & _
    ", " & Cant & ", #" & Me.fechaprocedencia & "#)"
'SubPorcionamiento
DoCmd.RunSQL "INSERT INTO SubPorcionamiento(Procedencia, CodProcesamiento, Cantidad, Fecha) values" & _
    " (" & Me.CodCotizacion & ", 0, " & Cant & ", #" & Me.fechaprocedencia & "#)"
Me.Requery
codsubpor = UltimoCodSubporcionamiento()
'Porcionado
DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia, CodSubPorcionamiento)" & _
    " values (" & PorcionDeCotizacion(Me.CodCotizacion) & ", 0, " & Cant & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & Me.CodCotizacion & ", " & codsubpor & ")"
DoCmd.SetWarnings True

Exit Sub
Salir:
MsgBox "No se ingresaron los datos correctamente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
