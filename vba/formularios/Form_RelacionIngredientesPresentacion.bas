' ==========================================================
' Modulo  : Form_RelacionIngredientesPresentacion
' Tipo    : 100  |  Lineas: 136
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.aproducto.Text = "" Then
    Me.Filter = "[Nombre] & ' ' & [CodIngrediente] & ' ' & [Unidad] like '*" & Me.aproducto.Text & "*'"
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

Private Sub buscargrupo_Click()
If IsNull(Me.atipo1) Then Me.atipo1 = ""
If IsNull(Me.atipo2) Then Me.atipo2 = ""
If IsNull(Me.acuenta) Then Me.acuenta = ""
If IsNull(Me.asubcuenta) Then Me.asubcuenta = ""

If Me.atipo1 = "" Then 'Buscamos por cuenta
    If Me.asubcuenta = "" And Me.acuenta <> "" Then 'Buscamos solo por cuenta
        Me.Filter = "IIf(IsNull([CodCeCoSubCuentas]),0,Dlookup('[CodCeCoCuentas]','[CeCoSubCuentas]','[CodCeCoSubCuentas]=' & [CodCeCoSubCuentas]))=" & Me.acuenta
        Me.FilterOn = True
    ElseIf Me.acuenta <> "" And Me.asubcuenta <> "" Then 'Buscamos solo por subcuenta
        Me.Filter = "[CodCeCoSubCuentas]=" & Me.asubcuenta
        Me.FilterOn = True
    Else 'vacio ambos
        Me.FilterOn = False
    End If
ElseIf Me.acuenta = "" Then  'Buscamos por tipo
    If Me.atipo2 = "" And Me.atipo1 <> "" Then  'Buscamos solo por tipo 1
        Me.Filter = "[TIPO2]='" & Me.atipo1 & "'"
        Me.FilterOn = True
    ElseIf Me.atipo1 <> "" And Me.atipo2 <> "" Then  'Buscamos solo por tipo 2 y tipo 1
        Me.Filter = "[Tipo]='" & Me.atipo2 & "' AND [TIPO2]='" & Me.atipo1 & "'"
        Me.FilterOn = True
    Else 'vacio ambos
        Me.FilterOn = False
    End If
Else 'No se busca nada , no hay nada seleccionado
    Me.FilterOn = False
End If

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


Private Sub asubcuenta_Exit(Cancel As Integer)
Me.atipo1 = ""
Me.atipo2 = ""
End Sub

Private Sub atipo1_Exit(Cancel As Integer)
Me.atipo2 = ""
Me.atipo2.Requery
End Sub

Private Sub atipo2_Exit(Cancel As Integer)
Me.acuenta = ""
Me.asubcuenta = ""
End Sub

Private Sub acuenta_Exit(Cancel As Integer)
Me.asubcuenta = ""
Me.asubcuenta.Requery
End Sub



Private Sub Comando563_Click()
DoCmd.OpenForm "RegistroIngredientesPresentacionNuevo"
[Forms]![RegistroIngredientesPresentacionNuevo]![metodo] = 1

[Forms]![RegistroIngredientesPresentacionNuevo]![conve] = 1
[Forms]![RegistroIngredientesPresentacionNuevo]![unco] = "Unid"
[Forms]![RegistroIngredientesPresentacionNuevo]![prese] = ""

[Forms]![RegistroIngredientesPresentacionNuevo]![conve].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![unco].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![prese].Enabled = False
End Sub

Private Sub Comando578_Click()
DoCmd.OpenForm "RegistroCeCoSubCuentas"
End Sub

Private Sub Comando584_Click()
DoCmd.OpenForm "RegistroIngredientesPresentacionNuevo"
[Forms]![RegistroIngredientesPresentacionNuevo]![metodo] = 2

[Forms]![RegistroIngredientesPresentacionNuevo]![codi] = Me.CodIngrediente
[Forms]![RegistroIngredientesPresentacionNuevo]![nomb] = Me.Nombre
[Forms]![RegistroIngredientesPresentacionNuevo]![unid] = Me.Unidad
[Forms]![RegistroIngredientesPresentacionNuevo]![tipoc] = Me.TIPO1
[Forms]![RegistroIngredientesPresentacionNuevo]![subc] = Me.CodCeCoSubCuentas
[Forms]![RegistroIngredientesPresentacionNuevo]![cont] = Me.CodControlAlmacenes

[Forms]![RegistroIngredientesPresentacionNuevo]![codi].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![nomb].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![unid].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![tipoc].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![subc].Enabled = False
[Forms]![RegistroIngredientesPresentacionNuevo]![cont].Enabled = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "Batidos Pitaya"
Me.ShortcutMenu = False
End Sub






