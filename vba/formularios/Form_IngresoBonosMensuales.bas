' ==========================================================
' Modulo  : Form_IngresoBonosMensuales
' Tipo    : 100
' Lineas  : 54
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database



Private Sub Comando198_Click()
Me.Requery
End Sub

Private Sub Comando754_Click()


On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer
Dim cot As Integer
Dim Cal As Double

miSQL = "SELECT DatosSistema.Nombre, DatosSistema.Ciudad, DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For I = 1 To canti

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO PorcentajeBonoSucursal(Sucursal, Porcentaje, Desde) values" & _
    " (" & rst("CodSistema") & ", 0, #" & Me.desde & "#)"
    DoCmd.SetWarnings True

    rst.MoveNext
Next I

rst.Close
MsgBox "Lista de Bonos creada correctamente de la quincena"
Me.Requery
Exit Sub

AgainAgain:
MsgBox "No se pudo generar Lista de Bonos correctamente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub intervalo_Change()
Me.desde.Requery
Me.hasta.Requery
End Sub
