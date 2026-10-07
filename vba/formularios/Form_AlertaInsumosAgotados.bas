' ==========================================================
' Modulo  : Form_AlertaInsumosAgotados
' Tipo    : 100
' Lineas  : 13
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database

Private Sub Comando458_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
Me.Filter = "[SI]+[INGRESO]-[CONSUMO]-[MERMA]<=[CONSUMOMAXIMODIARIO]*2 AND [SI]+[INGRESO]>0"
Me.FilterOn = True
End Sub
