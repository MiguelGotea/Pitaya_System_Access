' ==========================================================
' Modulo  : Form_SubPorcionamientovsPorcionamiento
' Tipo    : 100  |  Lineas: 62
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database



Private Sub Comando159_Click()
DoCmd.OpenForm "Porcionamiento"
[Forms]![Porcionamiento]![cing] = Me.CodIngrediente
[Forms]![Porcionamiento]![cfecha] = Me.Fecha
[Forms]![Porcionamiento]![csemana] = numerosemana(Me.Fecha)
[Forms]![Porcionamiento]![cproce] = Me.CodProcesamiento
[Forms]![Porcionamiento]![cprocedencia] = Me.Procedencia
[Forms]![Porcionamiento]![nprocedencia] = nombreproductocoti(Me.Procedencia)
[Forms]![Porcionamiento]![asubporcionamiento] = Me.CodSubPorcionamiento
[Forms]![Porcionamiento].Requery

End Sub

Private Sub Comando199_Click()
Me.Requery
End Sub


Public Sub Comando295_Click()
Me.Requery
Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

If codigoLocal() = 0 Then
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Reporte Semanal\" & Me.asemana
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Reporte Semanal"
Else
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
End If

If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Historial de Transformacion.pdf"
DoCmd.OutputTo acOutputForm, "SubPorcionamientovsPorcionamiento", acFormatPDF, archivo, False

DoCmd.Close acForm, "SubPorcionamientovsPorcionamiento"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)
Me.asemana = Me.semanaactual
Me.Requery
End Sub
