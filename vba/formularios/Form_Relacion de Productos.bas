' ==========================================================
' Modulo  : Form_Relacion de Productos
' Tipo    : 100
' Lineas  : 70
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.aproducto.Text = "" Then
    Me.Filter = "[Nombre] like '*" & Me.aproducto.Text & "*'"
    Me.FilterOn = True
    Me.aproducto.SelStart = Len(Me.aproducto)
Else
    Me.FilterOn = False
    Me.aproducto.SetFocus
End If

Exit Sub

NoResultados:
MsgBox "No se encontraron coincidencias"
Me.FilterOn = False
Me.aproducto.SetFocus
Me.aproducto.Text = Left(Me.aproducto.Text, Len(Me.aproducto.Text) - 1)
Call actualizarlista

End Sub




'Buscador de datos
Private Sub aproducto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
    Me.aproducto.Text = Left(Me.aproducto.Text, Len(Me.aproducto.Text) - 1)
    Call actualizarlista
Else
    Call actualizarlista
End If
End Sub

Private Sub CodigosBarra_DblClick(Cancel As Integer)
Dim codigob As String
codigob = InputBox("Codigo de barra para el producto " & Me.Nombre & " " & Me.Medida & ".", "Ingresar Codigo de Barra", "")

If codigob = "" Then
    MsgBox "no se registro ningun codigo de barra"
Else
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO CodigoBarraBatidos (CodBatido, CodigoBarra)" & _
    " VALUES ('" & Me.CodBatido & "', '" & codigob & "')"
    DoCmd.SetWarnings True
End If

Me.Requery
End Sub


Private Sub Comando355_Click()
DoCmd.OpenForm "HistorialdeProductosVenta"
[Forms]![HistorialdeProductosVenta]![cod] = [Forms]![Relacion de Productos]![CodBatido]
[Forms]![HistorialdeProductosVenta]![hgrupo] = [Forms]![Relacion de Productos]![CodGrupo]
Call Forms("[HistorialdeProductosVenta]").Comando57_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub





