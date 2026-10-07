' ==========================================================
' Modulo  : Form_Ingreso Inventario Mensual
' Tipo    : 100  |  Lineas: 60
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:11
' ==========================================================

Option Compare Database
Private Sub CrearListaUnidadesCotizacion(fech As Date)
'Lista completa ctizaciones *COntrol SI *NO Variables
'166 al 19/04/2019

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT Cotizaciones.CodCotizacion, DBIngredientes.TIPO1, DBIngredientes.Control" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (DBIngredientes.TIPO1<>'VARIABLES' AND DBIngredientes.Control<>0)"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Cotizacion](CodCotizacion, Fecha) values (" & rst("CodCotizacion") & ", #" & fech & "#)"
    DoCmd.SetWarnings True
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close
MsgBox "Se agregaron " & contador & " elementos"

Exit Sub

Nulo:
MsgBox "Error al crear datos"

End Sub

Private Sub Comando21_Click()
Me.fechaac = Me.fechaac - 1
Me.Requery
End Sub

Private Sub Comando23_Click()
Me.fechaac = Me.fechaac + 1
Me.Requery
End Sub

Private Sub Comando40_Click()
MsgBox "Esta a punto de crear la lista de productos de manera automatica"

If MsgBox("Desea generar lista completa de productos?", vbYesNo + vbQuestion) = vbYes Then
    Call CrearListaUnidadesCotizacion(Me.fechaac)
End If

Me.Requery
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
