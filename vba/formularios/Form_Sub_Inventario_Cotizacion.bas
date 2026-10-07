' ==========================================================
' Modulo  : Form_Sub_Inventario_Cotizacion
' Tipo    : 100
' Lineas  : 13
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Option Compare Database

Private Sub Cantidad_Click()
If [Forms]![Ingreso Inventario Almacen]![fechaac] < Date - 3 Then
    MsgBox "No se puede editar inventario de fechas pasadas"
    [Forms]![Ingreso Inventario Almacen]![fechaac].SetFocus
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
