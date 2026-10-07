' ==========================================================
' Modulo  : Form_PlanDespachoInsumosFIjosConsumibles
' Tipo    : 100  |  Lineas: 57
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database


Public Sub asucursal_Exit(Cancel As Integer)
If Not IsNull(asemana) Then

Me.seco1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.seco2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=4 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
End If
End Sub

Public Sub Comando1040_Click()
Me.Requery

Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

DoCmd.SelectObject acForm, "PlanDespachoInsumosFijosConsumibles", True
'DoCmd.RunCommand acCmdPrint

Dim Direccion, Direccion2 As String
Dim archivo As String
Dim archivo2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Despacho\" & Me.asemana & "\Detalle Sucursales"
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Despacho\" & Me.asemana
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.asucursal & " - Insumos Fijos Consumibles - Incremento " & Me.incre * 100 & "%.pdf"
DoCmd.OutputTo acOutputForm, "PlanDespachoInsumosFijosConsumibles", acFormatPDF, archivo, False

DoCmd.Close acForm, "PlanDespachoInsumosFijosConsumibles"
End Sub


Private Sub Comando91_Click()
Me.Requery
End Sub



Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub
