' ==========================================================
' Modulo  : Form_CompraInsumosFijosConsumiblesGlobal
' Tipo    : 100  |  Lineas: 62
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database



Public Sub asemana_Exit(Cancel As Integer)

Me.entregaleon = DLookup("[Carga]", "[PlanDespacho]", "[Local]=2 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.entregamatagalpa = DLookup("[Carga]", "[PlanDespacho]", "[Local]=4 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.entregaesteli = DLookup("[Carga]", "[PlanDespacho]", "[Local]=5 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.entregacolinas = DLookup("[Carga]", "[PlanDespacho]", "[Local]=11 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1

Me.entregavillafontana = DLookup("[Carga]", "[PlanDespacho]", "[Local]=9 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.entregaaltamira = DLookup("[Carga]", "[PlanDespacho]", "[Local]=7 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.entregagranada = DLookup("[Carga]", "[PlanDespacho]", "[Local]=10 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1
Me.entregamasaya = DLookup("[Carga]", "[PlanDespacho]", "[Local]=12 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1

Me.entregaproduccion = DLookup("[Carga]", "[PlanDespacho]", "[Local]=6 AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'") + 1

End Sub

Public Sub Comando199_Click()
Me.Requery

Dim Direccion, Direccion2 As String
Dim Direccion3 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\Plan de Compras\" & Me.asemana & "\Compras Quincenales"
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

archivo = Direccion & "\" & "Compra Quincenal Insumos Fijos Consumibles.pdf"
DoCmd.OutputTo acOutputForm, "CompraInsumosFijosConsumiblesGlobal", acFormatPDF, archivo, False

DoCmd.Close acForm, "CompraInsumosFijosConsumiblesGlobal"
End Sub

Private Sub Comando91_Click()
Me.Requery
End Sub


Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub
