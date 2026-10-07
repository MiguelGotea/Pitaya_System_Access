' ==========================================================
' Modulo  : Form_PlanProduccionMarcaPitaya
' Tipo    : 100
' Lineas  : 68
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database

Public Sub asemana_Exit(Cancel As Integer)
Me.ahabilitarleon = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarcolinas = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarnatura = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarmatagalpa = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitaresteli = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitaraltamira = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarvilla = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitargranada = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarmasaya = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.ahabilitarleon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarcolinas2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarnatura2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarmatagalpa2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitaresteli2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitaraltamira2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarvilla2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitargranada2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitarmasaya2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

End Sub

Public Sub Comando102_Click()
Me.Requery
End Sub

Public Sub Comando170_Click()
[Forms]![PlanProduccionMarcaPitaya].Form.Requery
[Forms]![PlanProduccionMarcaPitaya].Form.Printer.Orientation = acPRORLandscape
[Forms]![PlanProduccionMarcaPitaya].Form.Printer.LeftMargin = 0
[Forms]![PlanProduccionMarcaPitaya].Form.Printer.RightMargin = 0
[Forms]![PlanProduccionMarcaPitaya].Form.Printer.BottomMargin = 0
[Forms]![PlanProduccionMarcaPitaya].Form.Printer.TopMargin = 0

'DoCmd.SelectObject acForm, "PlanProduccionMarcaPitaya", True
'DoCmd.RunCommand acCmdPrint

Dim Direccion, Direccion2 As String
Dim archivo As String
Dim archivo2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Plan Semanal\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Plan Semanal\"
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Plan de Produccion Marca Pitaya.pdf"
DoCmd.OutputTo acOutputForm, "PlanProduccionMarcaPitaya", acFormatPDF, archivo, False

DoCmd.Close acForm, "PlanProduccionMarcaPitaya"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)

End Sub
