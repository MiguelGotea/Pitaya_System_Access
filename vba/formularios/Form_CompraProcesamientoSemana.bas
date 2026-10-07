' ==========================================================
' Modulo  : Form_CompraProcesamientoSemana
' Tipo    : 100
' Lineas  : 86
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database



Public Sub asemana_Exit(Cancel As Integer)
Dim leo1, leo2, colina1, colina2, natura1, natura2, mata1, mata2, este1, este2, villa1, villa2, alta1, alta2, grana1, grana2 As Double

leo1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
leo2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
colina1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
colina2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
natura1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
natura2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=13 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")


mata1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
mata2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
este1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
este2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

villa1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
villa2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
alta1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
alta2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
grana1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
grana2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
masa1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
masa2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")


Me.entregaleon = leo2 + 0.5
Me.entregacolinas = colina2 + 0.5
Me.entreganatura = natura2 + 0.5

Me.entregamatagalpa = mata2 + 0.5
Me.entregaesteli = este2 + 0.5

Me.entregavillafontana = villa2
Me.entregaaltamira = alta2
Me.entregagranada = grana2
Me.entregamasaya = masa2

End Sub

Public Sub Comando199_Click()
Me.Requery

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

archivo = Direccion & "\" & "Compra Semanal Insumos Congelados.pdf"
DoCmd.OutputTo acOutputForm, "CompraProcesamientoSemana", acFormatPDF, archivo, False

DoCmd.Close acForm, "CompraProcesamientoSemana"
End Sub

Private Sub Comando91_Click()
Me.Requery
End Sub


Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub
