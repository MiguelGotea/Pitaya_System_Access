' ==========================================================
' Modulo  : Form_Clientes Club Pitaya Interno
' Tipo    : 100  |  Lineas: 83
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database
Private Sub actualizarlista()
On Error GoTo NoResultados

If Not Me.bnombre.Text = "" Then
    Me.Filter = "[Nombre] like '*" & Me.bnombre.Text & "*' or [Apellidos] like '*" & Me.bnombre.Text & "*'"
    Me.FilterOn = True
    Me.bnombre.SelStart = Len(Me.bnombre)
Else
    Me.FilterOn = False
    Me.bnombre.SetFocus
End If

Exit Sub

NoResultados:
MsgBox "No se encontraron coincidencias"
Me.FilterOn = False
Me.bnombre.SetFocus
Me.bnombre.Text = Left(Me.bnombre.Text, Len(Me.bnombre.Text) - 1)
Call actualizarlista

End Sub

Private Sub bnombre_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 32 Then
    Me.bnombre.Text = Left(Me.bnombre.Text, Len(Me.bnombre.Text) - 1)
    Call actualizarlista
Else
    Call actualizarlista
End If

End Sub

Private Sub Comando178_Click()
Me.OrderBy = "[CodAfiliado] Asc"
Me.OrderByOn = True
End Sub

Private Sub Comando188_Click()
On Error GoTo Nulo
Dim strDefaultPrinter As String
Dim solonombre() As String
Dim soloapellido() As String

solonombre = Split(Me.Nombre)
soloapellido = Split(Me.Apellidos)
strDefaultPrinter = Application.Printer.DeviceName

DoCmd.OpenReport "StickerClientesClub", acViewReport
[Reports]![StickerClientesClub]![mcodigo] = Me.CodCliente
[Reports]![StickerClientesClub]![mnombre] = solonombre(0) & " " & soloapellido(0)
'DoCmd.OutputTo acOutputReport, "StickerClientesClub", acFormatPDF, "C:\Users\USER\Desktop\Sistema\sticker.pdf", False
DoCmd.SelectObject acReport, "StickerClientesClub"
Set Application.Printer = Application.Printers("HPRT D21")
DoCmd.PrintOut acSelection
Set Application.Printer = Application.Printers(strDefaultPrinter)
DoCmd.Close acReport, "StickerClientesClub"

Exit Sub
Nulo:
End Sub

Private Sub Comando87_Click()
Me.OrderBy = "[Nombre] Asc"
Me.OrderByOn = True
End Sub

Private Sub Comando88_Click()
Me.OrderBy = "[CodCliente] Asc"
Me.OrderByOn = True
End Sub

Private Sub Comando96_Click()
Me.OrderBy = "[Apellidos] Asc"
Me.OrderByOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
