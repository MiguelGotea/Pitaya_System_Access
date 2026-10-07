' ==========================================================
' Modulo  : Form_ComprasSemanalesSecos
' Tipo    : 100
' Lineas  : 89
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database


Public Sub asemana_Exit(Cancel As Integer)
    
Me.cleon1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cleon2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cnatura1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cnatura2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ccolinas1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ccolinas2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.cmatagalpa1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cmatagalpa2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cesteli1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cesteli2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.caltamira1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.caltamira2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cgranada1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cgranada2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cmasaya1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cmasaya2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cvilla1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.cvilla2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")


Me.cleon3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.ccolinas3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.cnatura3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1

Me.cmatagalpa3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.cesteli3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1

Me.caltamira3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.cgranada3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.cmasaya3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.cvilla3 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1

End Sub

Public Sub Comando102_Click()
Me.Requery
End Sub

Public Sub Comando170_Click()

Me.Requery
Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim Direccion3 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.asemana & "\Compras Semanales"
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.asemana
Direccion3 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras"

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

archivo = Direccion & "\" & "Compra Semanal Insumos Secos.pdf"
DoCmd.OutputTo acOutputForm, "ComprasSemanalesSecos", acFormatPDF, archivo, False

DoCmd.Close acForm, "ComprasSemanalesSecos"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)

End Sub
