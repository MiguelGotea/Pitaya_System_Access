' ==========================================================
' Modulo  : Form_Relacion de Productos Venta x Insumo
' Tipo    : 100  |  Lineas: 33
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================


Option Compare Database








Private Sub Comando354_Click()
Me.Requery
End Sub

Private Sub btnFiltrarVigentes_Click()
    If Me.Filter = "DBBatidos.Vigencia = True" And Me.FilterOn = True Then
        ' Si ya está filtrado, quitar el filtro (toggle)
        Me.FilterOn = False
        Me.btnFiltrarVigentes.Caption = "Solo Vigentes"
    Else
        ' Aplicar filtro
        Me.Filter = "DBBatidos.Vigencia = True"
        Me.FilterOn = True
        Me.btnFiltrarVigentes.Caption = "Ver Todos"
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub


