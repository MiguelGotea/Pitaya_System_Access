' ==========================================================
' Modulo  : Form_Resultados Promocion
' Tipo    : 100  |  Lineas: 110
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database
Private Sub ocultardatosdia()
Me.totalconpromo.Visible = False
Me.totalsinpromo.Visible = False
Me.totaldia.Visible = False
Me.ptotalconpromo.Visible = False
Me.ptotalsinpromo.Visible = False
End Sub






Private Sub Comando151_Click()
Me.Filter = ""
Me.FilterOn = False
Me.Filter = "semana = " & Me.ccsemana & " and CodPromocion = " & Me.bpromocion
Me.FilterOn = True
Me.Requery
Call ocultardatosdia
End Sub

Public Sub Comando278_Click()
If codigoLocal() = 0 Then 'Sistema local
    Exit Sub
End If

Me.Filter = ""
Me.FilterOn = False
Me.Filter = "semana = " & Me.asemana & " and CodPromocion <> 5"
Me.FilterOn = True
Me.Requery

Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

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

archivo = Direccion & "\Historial de Promociones.pdf"
DoCmd.OutputTo acOutputForm, "Resultados Promocion", acFormatPDF, archivo, False

Call ocultardatosdia
End Sub

Private Sub Comando366_Click()
Me.Filter = ""
Me.FilterOn = False
Me.Filter = "semana = " & Me.asemana & " and CodPromocion <> 5"
Me.FilterOn = True
Me.Requery

Call ocultardatosdia
End Sub

Private Sub Comando80_Click()

Dim dat As Date

Me.Filter = ""
Me.FilterOn = False
Me.Filter = "fecha >= #" & Me.bdesde & "# and fecha <= #" & Me.bhasta & "# and CodPromocion = " & Me.bpromocion
Me.FilterOn = True
Me.Requery

Me.totaldia = 0

For dat = Me.bdesde To Me.bhasta
    Me.totaldia = Me.totaldia + AcumuladoDiaGuardado(dat)
Next dat

Me.totalsinpromo.Visible = True
Me.totalconpromo.Visible = True
Me.totaldia.Visible = True
Me.ptotalsinpromo.Visible = True
Me.ptotalconpromo.Visible = True
End Sub

Private Sub DBBatidos_Nombre_DblClick(Cancel As Integer)
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & Me.CodPedido
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Filter = ""
Me.FilterOn = False
Me.Filter = "semana = " & Me.asemana & " and CodPromocion <> 5"
Me.FilterOn = True
Me.Requery
End Sub
