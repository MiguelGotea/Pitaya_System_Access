' ==========================================================
' Modulo  : Form_Filtro Ingresos Semana
' Tipo    : 100
' Lineas  : 74
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database

Private Sub bcodigo_Change()
If Not Me.bnombre = "" Then
    Me.bnombre = ""
End If
End Sub

Private Sub bnombre_KeyUp(KeyCode As Integer, Shift As Integer)
If Not Me.bcodigo = "" Then
    Me.bcodigo = ""
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
Else
    Me.Filter = "[CodCotizacion] = " & Me.bcodigo & " and [semana] >= " & numerosemana(Date) - Me.semanas
    Me.FilterOn = True
End If
End Sub

Private Sub Comando89_Click()
Me.bnombre = ""
Me.bcodigo = ""
Me.Filter = "[Control] = False and [semana] >= " & numerosemana(Date) - Me.semanas
Me.FilterOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 8000
'Deifnir origen de base de datos base de datos mixed
Me.RecordSource = "SELECT [DBIngredientes]![NombreSinProcesar] & ' ' & [Cotizaciones]![Marca] & ' ' & [Cotizaciones]![Linea] & ' ' & [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS Nombre," & _
" IngresosPitaya.Fecha, Cotizaciones.CodCotizacion, DBIngredientes.Control, [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad] AS Presentacion," & _
" IngresosPitaya.Cantidad, numerosemana([Fecha]) AS semana, IngresosPitaya.Expr1000" & _
" FROM IngresosPitaya INNER JOIN (DBIngredientes INNER JOIN Cotizaciones" & _
" ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" ORDER BY IngresosPitaya.Fecha DESC, [DBIngredientes]![NombreSinProcesar] & ' ' & [Cotizaciones]![Marca] & ' ' & [Cotizaciones]![Linea] & ' ' & [Cotizaciones]![Unidad] & ' ' & [Cotizaciones]![Capacidad]"
End Sub

Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.bnombre.Text = "" Then
    Me.Filter = "[Nombre] like '*" & Me.bnombre.Text & "*' and [semana] >= " & numerosemana(Date) - Me.semanas
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
