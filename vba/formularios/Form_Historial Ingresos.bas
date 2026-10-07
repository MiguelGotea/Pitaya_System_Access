' ==========================================================
' Modulo  : Form_Historial Ingresos
' Tipo    : 100  |  Lineas: 16
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando95_Click()
If MsgBox("Desea eliminar el ingreso de este item", vbYesNo, "Eliminar") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM IngresosPitaya WHERE CodIngresoPitaya = " & Me.CodIngresoPitaya
    DoCmd.SetWarnings True
    Me.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
