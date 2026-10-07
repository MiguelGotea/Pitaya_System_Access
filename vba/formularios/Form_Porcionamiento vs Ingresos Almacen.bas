' ==========================================================
' Modulo  : Form_Porcionamiento vs Ingresos Almacen
' Tipo    : 100  |  Lineas: 83
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Private Sub Comando108_Click()
DoCmd.OpenForm "Historial Porcionamiento Insumo"

[Forms]![Historial Porcionamiento Insumo]![cing] = Me.CodIngrediente
[Forms]![Historial Porcionamiento Insumo]![cprocedencia] = Me.CodCotizacion
[Forms]![Historial Porcionamiento Insumo]![nprocedencia] = nombreproductocoti(Me.CodCotizacion)
[Forms]![Historial Porcionamiento Insumo]![cproce] = 0 'COdigo de Procesamiento


If Me.asemana = numerosemana(Date) Then
    [Forms]![Historial Porcionamiento Insumo]![cfecha] = Date
    [Forms]![Historial Porcionamiento Insumo]![csemana] = numerosemana(Date)
    [Forms]![Historial Porcionamiento Insumo].Form.Filter = "[Procedencia]=" & Me.CodCotizacion & " AND [semana]=" & numerosemana(Date)
Else
    [Forms]![Historial Porcionamiento Insumo]![csemana] = Me.asemana
    [Forms]![Historial Porcionamiento Insumo].Form.Filter = "[Procedencia]=" & Me.CodCotizacion & " AND [semana]=" & Me.asemana
End If

[Forms]![Historial Porcionamiento Insumo].Form.FilterOn = True
[Forms]![Historial Porcionamiento Insumo].Requery
End Sub

Private Sub Comando114_Click()
DoCmd.OpenForm "Historial Ingresos", acNormal, , "[Semana]=" & Me.asemana & " AND [CodCotizacion]=" & Me.CodCotizacion & ""
End Sub



Private Sub Comando254_Click()
Me.Requery
End Sub

Public Sub Comando314_Click()
On Error GoTo CerrarYa

Me.Requery
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

archivo = Direccion & "\" & "Porcionamiento vs Ingresos Porciones.pdf"
DoCmd.OutputTo acOutputForm, "Porcionamiento vs Ingresos Almacen", acFormatPDF, archivo, False

DoCmd.Close acForm, "Porcionamiento vs Ingresos Almacen"
Exit Sub

CerrarYa:

DoCmd.CancelEvent
End Sub

Private Sub Comando94_Click()
DoCmd.OpenForm "Control Porcionamiento", acNormal, , "[CodIngrediente]='" & Me.CodIngrediente & "'"
[Forms]![Control Porcionamiento]![asemana] = Me.asemana
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)
End Sub
