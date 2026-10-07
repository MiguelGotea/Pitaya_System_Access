' ==========================================================
' Modulo  : Form_Ingreso Inventario Main
' Tipo    : 100  |  Lineas: 49
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database
Private Sub CrearListaUnidadesPorciones(fech As Date)
'Lista completa porciones almacen global

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT Cotizaciones.Marca, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones" & _
" WHERE Cotizaciones.Marca='Almacen Global'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [Inventario Main](CodCotizacion, Fecha) values (" & rst("CodCotizacion") & ", #" & fech & "#)"
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
Private Sub Comando40_Click()
MsgBox "Esta a punto de crear la lista de porciones de manera automatica"

If MsgBox("Desea generar lista completa de porciones?", vbYesNo + vbQuestion) = vbYes Then
    Call CrearListaUnidadesPorciones(Me.fechaac)
End If

Me.Requery
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
