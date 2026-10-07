' ==========================================================
' Modulo  : Form_RelacionOperariosGlobal
' Tipo    : 100
' Lineas  : 61
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.bnombre.Text = "" Then
    Me.Filter = "[Nombre] & ' ' & [Nombre2] & ' ' & [Apellido] & ' ' & [Apellido2] & ' ' & [CodOperario] like '*" & Me.bnombre.Text & "*'" & IIf(Me.afiltro = "Todos", "", IIf(Me.afiltro = "Activos", " And [Operativo]<>0", IIf(Me.afiltro = "Inactivos", " And [Operativo]=0", "")))
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
Me.OrderBy = "[CodOperario] Asc"
Me.OrderByOn = True
End Sub

Private Sub Comando96_Click()
Me.OrderBy = "[Apellido] Asc"
Me.OrderByOn = True
End Sub


