' ==========================================================
' Modulo  : Form_RegistrarInsumosFijosConsumibles
' Tipo    : 100  |  Lineas: 39
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()
On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad de productos a registrar", "Cantidad", 0)

DoCmd.SetWarnings False
If Me.adesdetabla = "[SubPreIngresosPitaya]" Or Me.adesdetabla = "[CambiosPreIngresosPitaya]" Then
    DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
    " values (" & Me.CodCotizacion & ", " & tota & ", " & Me.apreingre & ")"
Else
    DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, Fecha)" & _
    " values (" & Me.CodCotizacion & ", " & tota & ", #" & Me.fechaprocedencia & "#)"
End If
DoCmd.SetWarnings True

Me.icoti = 1
Me.guardado.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"
End Sub

Private Sub Form_Close()
If Me.adesdeform = "[Ingreso Inventario Pitaya]" Then
    Forms(Me.adesdeform).Form.Subformulario_Inventario_Cotizacion.Requery
Else
    Forms(Me.adesdeform).Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
