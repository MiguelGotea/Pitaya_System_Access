' ==========================================================
' Modulo  : Form_InventarioDisponiblePorcionar
' Tipo    : 100
' Lineas  : 69
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.aproducto.Text = "" Then
    Me.Filter = "[Nombre] like '*" & Me.aproducto.Text & "*' And [Descontinuado] =0"
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
Private Sub aproducto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
    Me.aproducto.Text = Left(Me.aproducto.Text, Len(Me.aproducto.Text) - 1)
    Call actualizarlista
Else
    Call actualizarlista
End If
End Sub

Private Sub asemana_Exit(Cancel As Integer)
Me.Requery
End Sub

Private Sub Comando16_Click()
DoCmd.OpenForm "Porcionamiento"

[Forms]![Porcionamiento]![cing] = Me.CodIngrediente
[Forms]![Porcionamiento]![cproce] = 0 'COdigo de Procesamiento
[Forms]![Porcionamiento]![cprocedencia] = Me.CodCotizacion
[Forms]![Porcionamiento]![nprocedencia] = nombreproductocoti(Me.CodCotizacion)
[Forms]![Porcionamiento]![aingrediente] = DLookup("[Nombre]", "DBIngredientes", "[CodIngrediente]='" & Me.CodIngrediente & "'")
[Forms]![Porcionamiento]![convproce] = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion]=" & Me.CodCotizacion & "")
[Forms]![Porcionamiento]![uniing] = DLookup("[Unidad]", "DBIngredientes", "[CodIngrediente]='" & Me.CodIngrediente & "'")
[Forms]![Porcionamiento]![unicoti] = DLookup("[Unidad]", "Cotizaciones", "[CodCotizacion]=" & Me.CodCotizacion & "")

If Me.asemana = numerosemana(Date) Then
    [Forms]![Porcionamiento]![cfecha] = Date
    [Forms]![Porcionamiento]![csemana] = numerosemana(Date)
    [Forms]![Porcionamiento].Form.Filter = "[Procedencia]=" & Me.CodCotizacion & " AND [semana]=" & numerosemana(Date)
Else
    [Forms]![Porcionamiento]![csemana] = Me.asemana
    [Forms]![Porcionamiento].Form.Filter = "[Procedencia]=" & Me.CodCotizacion & " AND [semana]=" & Me.asemana
End If

[Forms]![Porcionamiento].Form.FilterOn = True
[Forms]![Porcionamiento].Requery
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaactual = numerosemana(Date)
Me.asemana = numerosemana(Date)
Me.Requery
End Sub
