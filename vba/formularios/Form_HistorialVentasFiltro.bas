' ==========================================================
' Modulo  : Form_HistorialVentasFiltro
' Tipo    : 100  |  Lineas: 55
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database



Private Sub Comando102_Click()
Me.Filter = "numerosemana([Fecha])=" & Me.asemana & " AND [POS]<>0 AND ([Delivery]=5 Or [Delivery]=6 Or [Delivery]=8)"
Me.FilterOn = True
Me.Requery

End Sub

Private Sub Comando107_Click()
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & Me.CodPedido
End Sub

Private Sub Comando133_Click()

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


Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

archivo = Direccion & "\" & Me.tituloprincipal.Caption & ".pdf"
DoCmd.OutputTo acOutputForm, "HistorialVentasFiltro", acFormatPDF, archivo, False


DoCmd.Close acForm, "HistorialVentasFiltro"
End Sub

Private Sub Form_Close()
If CurrentProject.AllForms("Cierre por Turno 3").IsLoaded Then
    Call Forms("Cierre por Turno 3").actualizariconoscierre
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
