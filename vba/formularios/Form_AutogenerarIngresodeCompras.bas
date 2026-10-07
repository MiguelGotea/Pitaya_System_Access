' ==========================================================
' Modulo  : Form_AutogenerarIngresodeCompras
' Tipo    : 100
' Lineas  : 41
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database

Private Sub fechaac_Change()
Me.Requery

End Sub

Private Sub Comando120_Click()
'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT Compras.Fecha, Compras.CodCotizacion, Compras.Cantidad, Compras.NumeroFactura" & _
" FROM Compras WHERE (Compras.NumeroFactura)='" & Me.NumeroFactura & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
Do While Not rst.EOF
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO IngresosPitaya(CodCotizacion, Cantidad, Fecha) values" & _
    " (" & rst("CodCotizacion") & ", " & rst("Cantidad") & ", #" & rst("Fecha") & "#)"
    DoCmd.SetWarnings True
    contador = contador + 1
    
    rst.MoveNext
Loop
rst.Close
MsgBox "Se agregaron " & contador & " productos al registro de ingresos"

Exit Sub

Nulo:
MsgBox "Error al crear datos"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub
