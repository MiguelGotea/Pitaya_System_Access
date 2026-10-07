' ==========================================================
' Modulo  : Form_VentasCodigoBarras
' Tipo    : 100
' Lineas  : 49
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database

Private Sub codigo_Exit(Cancel As Integer)
On Error GoTo Salir
Dim val As Long
Dim codin As Long
Dim cantin As Long
Dim codinglob As Long
Dim codpro As String

codpro = DLookup("[CodBatido]", "[CodigoBarraBatidos]", "[CodigoBarra]='" & Me.codigo & "'")
Call IngresarPedido(codpro, Me.anotapedido, 0)



Me.acodigo = codpro
Me.acantidad = 1
Me.aprodu = DLookup("[Nombre] & ' ' & [Medida]", "[DBBatidos]", "[CodBatido]='" & codpro & "'")
    
Me.codigo = ""
Me.codigo.SetFocus

Exit Sub

Salir:
If IsNull(Me.codigo) = True Or Me.codigo = "" Then
Else
    MsgBox "Codigo de Barras no existe"
End If

Me.codigo = ""
Me.codigo.SetFocus

'Me.Codigo.SelStart = 1
'SendKeys "{TAB}"
End Sub



Private Sub Form_Close()
If CurrentProject.AllForms("Nota de Pedido").IsLoaded Then
    [Forms]![Nota de Pedido].Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
