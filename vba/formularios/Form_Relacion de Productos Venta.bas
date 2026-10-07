' ==========================================================
' Modulo  : Form_Relacion de Productos Venta
' Tipo    : 100  |  Lineas: 70
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:08
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
'******************************** INICIO DE FORMULAS DE BOTONES Y CONTROLES ******************
'
'Borrar datos cuando se selecciona un tipo de buscador

Private Sub asubgrupo_GotFocus()
Me.aproducto.Value = ""
Me.imagen.Picture = ""
End Sub

Private Sub aproducto_GotFocus()
Me.asubgrupo.Value = ""
Me.imagen.Picture = ""
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

Private Sub asubgrupo_Change()
On Error Resume Next
Me.Filter = "[CodGrupo] = " & Me.asubgrupo
Me.FilterOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub


'*************************** FOTO DE PRODUCTO ****************
Private Sub NombreSinProcesar_GotFocus()
If Not Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Ventas\" & Me.CodBatido & ".png") = "" Then
    Me.imagen.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Ventas\" & Me.CodBatido & ".png"
Else
    Me.imagen.Picture = ""
End If
End Sub


