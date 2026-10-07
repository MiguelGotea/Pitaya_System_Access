' ==========================================================
' Modulo  : Form_HistorialdeProductosVenta
' Tipo    : 100  |  Lineas: 29
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database



Private Sub Comando513_Click()
Dim menu As String
menu = "Menu PITAYA Global"
DoCmd.OpenForm "Menu PITAYA Global"
End Sub

Public Sub Comando57_Click()
Me.acodbatido = Me.cod
Me.anombre = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.amedida = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.aprecio = DLookup("[Precio]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.acodgrupo = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.avigencia = DLookup("[Vigencia]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.Requery
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub hgrupo_Exit(Cancel As Integer)
Me.cod.Requery
End Sub
