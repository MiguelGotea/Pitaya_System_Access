' ==========================================================
' Modulo  : Form_RequerimientoPorcionesSemana
' Tipo    : 100  |  Lineas: 49
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database



Private Sub Comando199_Click()
Me.Etiqueta190.Visible = False
Me.rango.Visible = False

Me.Requery
DoCmd.SelectObject acForm, "RequerimientoPorcionesSemana", True
DoCmd.RunCommand acCmdPrint

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Compras"
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Plan Porciones 4(" & Me.entregamatagalpa & ") 2(" & Me.entregaleon & ") 5(" & Me.entregaesteli & ") 7(" & Me.entregamanagua & ").pdf"
DoCmd.OutputTo acOutputForm, "RequerimientoPorcionesSemana", acFormatPDF, archivo, False
End Sub

Private Sub Comando91_Click()
Me.Requery
End Sub





Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
End Sub

Private Sub Form_Open(Cancel As Integer)
'Me.envio = DLookup("[IngresoInsumos]", "DatosSistema")
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

Me.Requery
End Sub
