' ==========================================================
' Modulo  : Form_RevisionInventariosCargadosSucursalesFijos
' Tipo    : 100  |  Lineas: 18
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database





Private Sub Comando91_Click()
Me.Requery
End Sub



Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
