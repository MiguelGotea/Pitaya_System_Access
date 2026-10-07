' ==========================================================
' Modulo  : Form_SeguimientoPlanPorciones
' Tipo    : 100
' Lineas  : 62
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database


Public Sub asucursal_Exit(Cancel As Integer)
Me.seco1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.seco2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=1 AND [CodAlmacenamiento]=3 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.refri1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=1 AND [CodAlmacenamiento]=2 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.refri2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=1 AND [CodAlmacenamiento]=2 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.conge1 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=1 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
Me.conge2 = DLookup("[Carga]", "[PlanDespacho]", "[Local]=" & Me.asucursal & " AND [Lista]=1 AND [CodAlmacenamiento]=1 AND [Fecha]=2 AND [SemanaMes]= '" & paroimpar(Me.asemana) & "'")
End Sub

Public Sub Comando1040_Click()
Me.Requery
Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

DoCmd.SelectObject acForm, "SeguimientoPlanPorciones", True
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

archivo = Direccion & "\" & Me.asucursal & " - Productos Porciones Totales - Incremento " & Me.incre * 100 & "%.pdf"
DoCmd.OutputTo acOutputForm, "SeguimientoPlanPorciones", acFormatPDF, archivo, False

DoCmd.Close acForm, "SeguimientoPlanPorciones"
End Sub


Private Sub Comando91_Click()
Me.Requery
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
If CurrentProject.AllForms("RegistroPreIngresosPitaya").IsLoaded Then
    [Forms]![RegistroPreingresosPitaya].Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
