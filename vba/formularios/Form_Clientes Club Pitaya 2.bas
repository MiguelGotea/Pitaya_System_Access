' ==========================================================
' Modulo  : Form_Clientes Club Pitaya 2
' Tipo    : 100
' Lineas  : 77
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database



Private Sub Comando207_Click()
Call DescargarTablaCompleta("VentasGlobalesAccessCSV", "VentasGlobalesAccessCSVFiltradoCliente", "CodCliente = " & Me.membresia & " AND local <> " & codigoLocal())
Dim clientebuscar As Variant
clientebuscar = DatosClienteClubGlobal(Me.membresia)
Call CargarVentasInternoExternoClienteHaciaTabla(Me.membresia) 'crear tabla de istorial local y externo

DoCmd.OpenForm "Historial Cliente 2"
[Forms]![Historial Cliente 2]![acodigo] = Me.membresia
[Forms]![Historial Cliente 2]![anombre] = clientebuscar(1)
[Forms]![Historial Cliente 2]![apuntos] = clientebuscar(0)
[Forms]![Historial Cliente 2]![iniciales] = clientebuscar(3)
[Forms]![Historial Cliente 2].Form.Requery

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.bnombre.Text = "" Then
    Me.Filter = "[Nombre] & ' ' & [Apellido] & ' ' & [membresia] like '*" & Me.bnombre.Text & "*'"
    Me.FilterOn = True
    Me.bnombre.SelStart = Len(Me.bnombre)
Else
    Me.FilterOn = False
    Me.bnombre.SetFocus
End If

Exit Sub

NoResultados:
MsgBox "No se encontraron coincidencias"
Me.FilterOn = False
Me.bnombre.SetFocus
Me.bnombre.Text = Left(Me.bnombre.Text, Len(Me.bnombre.Text) - 1)
Call actualizarlista

End Sub

Private Sub bnombre_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 32 Then
    Me.bnombre.Text = Left(Me.bnombre.Text, Len(Me.bnombre.Text) - 1)
    Call actualizarlista
Else
    Call actualizarlista
End If

End Sub



Private Sub Comando87_Click()
Me.OrderBy = "[Nombre] Asc"
Me.OrderByOn = True
End Sub

Private Sub Comando88_Click()
Me.OrderBy = "[membresia] Asc"
Me.OrderByOn = True
End Sub

Private Sub Comando96_Click()
Me.OrderBy = "[Apellido] Asc"
Me.OrderByOn = True
End Sub


