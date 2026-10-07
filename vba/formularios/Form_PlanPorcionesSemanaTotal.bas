' ==========================================================
' Modulo  : Form_PlanPorcionesSemanaTotal
' Tipo    : 100  |  Lineas: 152
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database


Public Sub asemana_Exit(Cancel As Integer)
Me.secoleon1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secoleon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladoleon1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladoleon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secovilla1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secovilla2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladovilla1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladovilla2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secoMASAYA1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secomasaya2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladomasaya1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladomasaya2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secomata1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secomata2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladomata1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladomata2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secoesteli1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secoesteli2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladoesteli1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladoesteli2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secoaltamira1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secoaltamira2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladoaltamira1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladoaltamira2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secocolinas1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secocolinas2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladocolinas1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladocolinas2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.seconatura1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.seconatura2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladonatura1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladonatura2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.secogranada1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.secogranada2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladogranada1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.congeladogranada2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Dim templeon3 As Double
Dim tempcolinas3 As Double
Dim tempnatura3 As Double
Dim tempmata3 As Double
Dim tempesteli3 As Double

templeon3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana + 1) & "'")
Me.congeladoleon3 = IIf(Me.congeladoleon2 > templeon3 + 1, Me.congeladoleon2, templeon3 + 1)

tempcolinas3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana + 1) & "'")
Me.congeladocolinas3 = IIf(Me.congeladocolinas2 > tempcolinas3 + 1, Me.congeladocolinas2, tempcolinas3 + 1)

tempnatura3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana + 1) & "'")
Me.congeladonatura3 = IIf(Me.congeladonatura2 > tempnatura3 + 1, Me.congeladonatura2, tempnatura3 + 1)

tempmata3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana + 1) & "'")
Me.congeladomata3 = IIf(Me.congeladomata2 > tempmata3 + 1, Me.congeladomata2, tempmata3 + 1)

tempesteli3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana + 1) & "'")
Me.congeladoesteli3 = IIf(Me.congeladoesteli2 > tempesteli3 + 1, Me.congeladoesteli2, tempesteli3 + 1)

Me.Requery
End Sub

Public Sub Comando1040_Click()

Dim Direccion, Direccion2 As String
Dim archivo As String
Dim archivo2 As String
Dim archivo3 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Plan Semanal\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Plan Semanal"
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion"

If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If


Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

archivo = Direccion & "\" & "Plan de Produccion Porciones.pdf"
DoCmd.OutputTo acOutputForm, "PlanPorcionesSemanaTotal", acFormatPDF, archivo, False

DoCmd.Close acForm, "PlanPorcionesSemanaTotal"
End Sub

Private Sub Comando91_Click()
Me.Requery
End Sub


Private Sub Comando951_Click()

Me.Requery

Dim Direccion, Direccion2 As String
Dim Direccion3 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Plan Semanal\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion\Plan Semanal"
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Produccion"

If Dir(Direccion3, vbDirectory) = "" Then
    MkDir Direccion3
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Plan de Produccion Semanal.pdf"
DoCmd.OutputTo acOutputForm, "PlanPorcionesSemanaTotal", acFormatPDF, archivo, False
End Sub



Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
