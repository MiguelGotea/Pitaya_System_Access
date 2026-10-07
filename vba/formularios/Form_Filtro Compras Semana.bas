' ==========================================================
' Modulo  : Form_Filtro Compras Semana
' Tipo    : 100
' Lineas  : 134
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:17
' ==========================================================
Option Compare Database

Private Sub bcodigo_Change()
If Not Me.bnombre = "" Then
    Me.bnombre = ""
End If
If Not Me.cproovedor = "" Then
    Me.cproovedor = ""
End If
End Sub

Private Sub cproovedor_Change()
If Not Me.bnombre = "" Then
    Me.bnombre = ""
End If
If Not Me.bcodigo = "" Then
    Me.bcodigo = ""
End If
End Sub

Private Sub bnombre_KeyUp(KeyCode As Integer, Shift As Integer)
If Not Me.bcodigo = "" Then
    Me.bcodigo = ""
End If
If Not Me.cproovedor = "" Then
    Me.cproovedor = ""
End If

If KeyCode = 32 Then
    Me.bnombre.Text = Left(Me.bnombre.Text, Len(Me.bnombre.Text) - 1)
    Call actualizarlista
Else
    Call actualizarlista
End If
End Sub

Private Sub Comando133_Click()
If Not Me.bnombre = "" Then
    Me.bnombre.SetFocus
    Call actualizarlista
ElseIf Not Me.bcodigo = "" Then
    Me.Filter = "[CodCotizacion] = " & Me.bcodigo & " and [semana] >= " & numerosemana(Date) - Me.semanas
    Me.FilterOn = True
Else
    Me.Filter = "[CodProveedor] = " & Me.cproovedor & " and [semana] >= " & numerosemana(Date) - Me.semanas
    Me.FilterOn = True
End If
End Sub

Private Sub Comando152_Click()
Me.bcodigo = ""
Me.bnombre = ""
Me.Filter = "[Tipo] <> 'CAJA' and month([Fecha]) = " & Me.bmes & " and year([Fecha]) = " & Me.bano
Me.FilterOn = True
End Sub

Private Sub Comando159_Click()
MsgBox "Lista de Tipos de Pago:" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
       "TIPO 1: Efectivo" & Chr(10) & Chr(13) & _
       "TIPO 3: Cuenta Pagos Corriente" & Chr(10) & Chr(13) & _
       "TIPO 4: Cuenta Depositos Dolares" & Chr(10) & Chr(13) & _
       "TIPO 5: Cuenta Depositos Cordobas" & Chr(10) & Chr(13) & _
       "TIPO 6: Credito Visa Connect Miles" & Chr(10) & Chr(13) & _
       "TIPO 7: Credito Ameican Express Pricesmart"
End Sub

Private Sub Comando234_Click()
Me.bcodigo = ""
Me.bnombre = ""
Me.Filter = "[Tipo] = 'CAJA' and month([Fecha]) = " & Me.bmes & " and year([Fecha]) = " & Me.bano & " and [local] = 2"
Me.FilterOn = True
End Sub

Private Sub Comando236_Click()
DoCmd.OpenForm "Lista de Proovedores", acNormal
End Sub

Private Sub Comando248_Click()
Me.bcodigo = ""
Me.bnombre = ""
Me.Filter = "[Tipo] = 'CAJA' and month([Fecha]) = " & Me.bmes & " and year([Fecha]) = " & Me.bano & " and [local] = 4"
Me.FilterOn = True
End Sub

Private Sub Comando258_Click()
Me.bcodigo = ""
Me.bnombre = ""

Me.Filter = "[Tipo] = '" & Me.atipo & "' and month([Pagado]) = " & Me.bmes & " and year([Pagado]) = " & Me.bano
Me.FilterOn = True
Me.OrderBy = "[Pagado] DESC, [CodProveedor] DESC, [Destino] DESC, [CodCotizacion], [Pagado], [Destino]"
Me.OrderByOn = True


End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

Me.InsideHeight = 8000
'Deifnir origen de base de datos base de datos mixed
Me.RecordSource = "SELECT Compras.CodCotizacion, Compras.CodProveedor, Compras.local, Compras.Pagado, Compras.Fecha, Compras.Destino," & _
" [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS Presentacion, Compras.CodIngresoAlmacen, Compras.Cantidad," & _
" Compras.CostoTotal, Compras.Observaciones, numerosemana([Compras]![Fecha]) AS semana, DBIngredientes.TIPO2, DBIngredientes.TIPO1," & _
" Compras.Tipo, Proovedores.Nombre, [DBIngredientes]![NombreSinProcesar] & ' ' & [Cotizaciones]![Marca] & ' ' & [Cotizaciones]![Linea] & ' ' & [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS NombreProducto" & _
" FROM (DBIngredientes INNER JOIN (Cotizaciones INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion)" & _
" ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN Proovedores ON Compras.CodProveedor = Proovedores.CodProovedor IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" ORDER BY Compras.Fecha DESC, Compras.CodProveedor DESC, Compras.Destino DESC, Compras.CodCotizacion, Compras.Pagado, Compras.Destino"
End Sub


Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.bnombre.Text = "" Then
    Me.Filter = "[NombreProducto] like '*" & Me.bnombre.Text & "*' and [semana] >= " & numerosemana(Date) - Me.semanas
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
