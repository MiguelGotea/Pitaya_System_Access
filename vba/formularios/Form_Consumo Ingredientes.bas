' ==========================================================
' Modulo  : Form_Consumo Ingredientes
' Tipo    : 100
' Lineas  : 52
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando15_Click()
Me.Requery
End Sub

Private Sub Comando250_Click()
Me.asemana = Me.asemana - 1
Me.Requery
End Sub

Private Sub Comando253_Click()
Me.asemana = Me.asemana + 1
Me.Requery
End Sub

Private Sub Comando565_Click()
DoCmd.OpenForm "Control Porcionamiento", acNormal, , "[CodIngrediente]='" & Me.aingrediente & "'"
[Forms]![Control Porcionamiento]![asemana] = Me.asemana
End Sub

Private Sub Comando591_Click()
DoCmd.OpenReport "KardexIngredienteSemana", acViewReport
[Reports]![KardexIngredienteSemana].Report.OrderBy = "Fecha Asc, tipo Asc"
[Reports]![KardexIngredienteSemana].Report.OrderByOn = True


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

archivo = Direccion & "\Kardex - " & Me.aingrediente & ".pdf"
DoCmd.OutputTo acOutputReport, "KardexIngredienteSemana", acFormatPDF, archivo, False

DoCmd.Close acReport, "KardexIngredienteSemana"



End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
