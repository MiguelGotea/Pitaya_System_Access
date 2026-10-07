' ==========================================================
' Modulo  : Form_ConsumoPitayaStore
' Tipo    : 100  |  Lineas: 59
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:08
' ==========================================================

Option Compare Database



Public Sub alocali_Exit(Cancel As Integer)
Me.ahabilitar = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.alocali & " AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitar2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.alocali & " AND [Lista]=3 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")

Me.ahabilitarnp = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.alocali & " AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.ahabilitar2np = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.alocali & " AND [Lista]=3 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
End Sub

Public Sub Comando102_Click()
Me.Requery
End Sub



Public Sub Comando170_Click()
Me.Requery

Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

DoCmd.SelectObject acForm, "ConsumoPitayaStore", True
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

archivo = Direccion & "\" & Me.alocali & " - Productos de Mostrador - Incremento " & Me.incre * 100 & "%.pdf"
DoCmd.OutputTo acOutputForm, "ConsumoPitayaStore", acFormatPDF, archivo, False

DoCmd.Close acForm, "ConsumoPitayaStore"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaactual = numerosemana(Date)


End Sub
