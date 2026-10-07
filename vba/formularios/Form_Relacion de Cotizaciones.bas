' ==========================================================
' Modulo  : Form_Relacion de Cotizaciones
' Tipo    : 100
' Lineas  : 119
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.aproducto.Text = "" Then
    Me.Filter = "[NombreSinProcesar] & ' ' & [CodCotizacion] like '*" & Me.aproducto.Text & "*'"
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
Private Sub agrupo_GotFocus()
Me.aproducto.Value = ""
Me.imagen.Picture = ""
'Me.FilterOn = False
End Sub

Private Sub asubgrupo_GotFocus()
Me.aproducto.Value = ""
Me.imagen.Picture = ""
'Me.FilterOn = False
End Sub

Private Sub aproducto_GotFocus()
Me.asubgrupo.Value = ""
Me.agrupo.Value = ""
Me.imagen.Picture = ""
'Me.FilterOn = False
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

Private Sub agrupo_Change()
Me.Filter = "[TIPO2] = '" & Me.agrupo & "'"
Me.FilterOn = True
Me.asubgrupo.Value = ""
Me.asubgrupo.Requery
Me.asubgrupo.SetFocus
End Sub

Private Sub asubgrupo_Change()
Me.Filter = "[Tipo] = '" & Me.asubgrupo & "' and [TIPO2] = '" & Me.agrupo & "'"
Me.FilterOn = True
End Sub

Private Sub buscargrupo_Click()
Me.Filter = "[Tipo] = '" & Me.asubgrupo & "' and [TIPO2] = '" & Me.agrupo & "'"
Me.FilterOn = True
End Sub

Private Sub CodigosBarra_DblClick(Cancel As Integer)
Dim codigob As String
codigob = InputBox("Codigo de barra para el producto " & nombreproductocoti(Me.CodCotizacion) & ".", "Ingresar Codigo de Barra", "")

If codigob = "" Then
    MsgBox "no se registro ningun codigo de barra"
Else
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO CodigoBarraCotizacion (CodCotizacion, CodigoBarra)" & _
    " VALUES (" & Me.CodCotizacion & ", '" & codigob & "')"
    DoCmd.SetWarnings True
End If

Me.Requery
End Sub



Private Sub compradirectasucursal_DblClick(Cancel As Integer)
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "desbloquearcompralocal"
[Forms]![IngresoClavePrivado].variable = Me.CodCotizacion
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub NombreSinProcesar_DblClick(Cancel As Integer)

DoCmd.OpenForm "Analisis de Productos"
[Forms]![Analisis de Productos]![CodigoBusqueda] = [Forms]![Relacion de Cotizaciones]![CodCotizacion]
[Forms]![Analisis de Productos]![aingrediente] = [Forms]![Relacion de Cotizaciones]![CodIngrediente]
End Sub

'*************************** FOTO DE PRODUCTO ****************
Private Sub NombreSinProcesar_GotFocus()
If Not Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Productos\" & Me.CodCotizacion & ".png") = "" Then
    Me.imagen.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Productos\" & Me.CodCotizacion & ".png"
    Else
    Me.imagen.Picture = ""
End If
End Sub


