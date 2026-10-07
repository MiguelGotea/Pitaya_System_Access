' ==========================================================
' Modulo  : Form_HistoricoConsumoPorcionesDescargado
' Tipo    : 100  |  Lineas: 58
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database


Public Sub Comando170_Click()
'Call actualizarfiltroporcionesvigentes
Call Comando91_Click
'On Error GoTo Error

Dim Direccion, Direccion2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

If codigoLocal() = 0 Then
    DoCmd.OutputTo acOutputQuery, "HistoricoConsumoPorcionesDescargadoCentral", acFormatXLSX, Direccion & "\" & nombrelocalglobal(Me.asucursal) & " - " & Me.semanaactual & " hasta " & Me.semanaactual - 3 & " Porciones.xlsx", False
Else
    DoCmd.OutputTo acOutputQuery, "HistoricoConsumoPorcionesDescargado", acFormatXLSX, Direccion & "\" & nombrelocalglobal(Me.asucursal) & " - " & Me.semanaactual & " hasta " & Me.semanaactual - 3 & " Porciones.xlsx", False
End If

MsgBox "Excel descargado correctamente en la siguiente ruta \Google Drive BP\Datos Descargados de Sistema\Proyeccion Consumos"
Exit Sub
Error:
MsgBox "No se exporto el excel correctamente"

End Sub

Private Sub Comando91_Click()
If Me.asucursal <> 0 Then

    Dim rst As DAO.Recordset
    Dim miSQL As String
    miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
    " FROM StatusSucursales WHERE (StatusSucursales.Sucursal<>0 AND StatusSucursales.Activo<>0)"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    Do While Not rst.EOF
        Call importartablaespecifica("Pitaya" & rst("CodLocal") & "_DB", "CalConsumoNoPorcionSemanal", "CalConsumoNoPorcionSemanal" & rst("CodLocal"), 3)
        Call importartablaespecifica("Pitaya" & rst("CodLocal") & "_DB", "CalConsumoPorcionSemanal", "CalConsumoPorcionSemanal" & rst("CodLocal"), 3)
        rst.MoveNext
    Loop
    rst.Close

Else
    Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "CalConsumoNoPorcionSemanal", "CalConsumoNoPorcionSemanal", 1)
    Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "CalConsumoPorcionSemanal", "CalConsumoPorcionSemanal", 1)
End If
Me.Requery
End Sub



Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
