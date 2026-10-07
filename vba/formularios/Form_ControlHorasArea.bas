' ==========================================================
' Modulo  : Form_ControlHorasArea
' Tipo    : 100  |  Lineas: 29
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Private Sub Comando198_Click()
If IsNull(Me.CodigoBusqueda) = True Then
    Me.Filter = "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "#"
    Me.FilterOn = True
Else
    Me.Filter = "fecha >= #" & Me.desde & "# and fecha <= #" & Me.hasta & "# and [CodOperario]= " & Me.CodigoBusqueda
    Me.FilterOn = True
End If
End Sub





Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Call importartablasweb
End Sub

Private Sub intervalo_Change()
Me.desde.Requery
Me.hasta.Requery

End Sub


