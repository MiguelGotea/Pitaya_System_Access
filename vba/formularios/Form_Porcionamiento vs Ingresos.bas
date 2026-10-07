' ==========================================================
' Modulo  : Form_Porcionamiento vs Ingresos
' Tipo    : 100  |  Lineas: 89
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:11
' ==========================================================

Option Compare Database

Private Sub Comando108_Click()
DoCmd.OpenForm "Porcionamiento"

[Forms]![Porcionamiento]![cing] = Me.CodIngrediente
[Forms]![Porcionamiento]![aingrediente] = DLookup("[Nombre]", "DBIngredientes", "[CodIngrediente]='" & Me.CodIngrediente & "'")
[Forms]![Porcionamiento]![cprocedencia] = Me.CodCotizacion
[Forms]![Porcionamiento]![nprocedencia] = nombreproductocoti(Me.CodCotizacion)
[Forms]![Porcionamiento]![cproce] = 0 'COdigo de Procesamiento
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





Private Sub Comando254_Click()
Me.Requery
Me.Filter = "[SI] <> 0 or [Ingresos] <> 0 or [porcionado] <> 0 or [SFREAL] <> 0"
Me.FilterOn = True
Me.aproducto = ""

End Sub

Public Sub Comando314_Click()
Me.Requery
Me.Filter = "[SI] <> 0 or [Ingresos] <> 0 or [porcionado] <> 0 or [SFREAL] <> 0"
Me.FilterOn = True
Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Porcionamiento vs Ingresos.pdf"
DoCmd.OutputTo acOutputForm, "Porcionamiento vs Ingresos", acFormatPDF, archivo, False

DoCmd.Close acForm, "Porcionamiento vs Ingresos"
End Sub

Private Sub Comando333_Click()
If IsNull(Me.asemana) = True Or Me.asemana = "" Then
    MsgBox "Ingresar numero de semana"
Else
    Me.Filter = "[CodIngrediente] = '" & Me.aproducto & "'"
    Me.FilterOn = True
    Me.Requery
End If
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)
Me.aproducto = ""
End Sub
