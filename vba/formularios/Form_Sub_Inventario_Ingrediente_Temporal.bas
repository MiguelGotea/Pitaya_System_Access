' ==========================================================
' Modulo  : Form_Sub_Inventario_Ingrediente_Temporal
' Tipo    : 100  |  Lineas: 26
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()

On Error GoTo AgainAgain
Dim tota As Double

tota = InputBox("Ingrese la cantidad a registrar del siguiente producto: " & UCase(Me.Nombre), "Cantidad", 0)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO [Inventario Ingrediente Temporal](CodIngrediente, Cantidad, Fecha, CodOperario)" & _
" values ('" & Me.CodIngrediente & "', " & tota - StockIngrediente(Me.CodIngrediente, numerosemana([Forms]![Ingreso Inventario Pitaya]![fechaac]) + 1) & ", #" & [Forms]![Ingreso Inventario Pitaya]![fechaac] & "#, " & [Forms]![Ingreso Inventario Pitaya]![codigologin] & ")"
DoCmd.SetWarnings True

Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se registro correctamente, ingresarlo otra vez"

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
