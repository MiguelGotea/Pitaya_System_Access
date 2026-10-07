' ==========================================================
' Modulo  : Form_ChatSistema
' Tipo    : 100  |  Lineas: 115
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:08
' ==========================================================

Option Compare Database

Private Sub actualizarlista()

Dim cantactual As Integer
Dim fondochat As String

Me.Fecha1.Visible = False
Me.Fecha2.Visible = False
Me.Lado1.Visible = False
Me.Lado2.Visible = False
Me.RecordSource = ""
If DCount("*", "MSysObjects", "Type=1 AND Name='TempMensajes'") > 0 Then
DoCmd.RunSQL "DROP TABLE TempMensajes"
End If

'Crear carpeta temporal
DoCmd.SetWarnings False
DoCmd.RunSQL "SELECT x.* INTO TempMensajes" & _
" FROM (SELECT *," & Me.areceptor & " as Emisor FROM ChatPitaya IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & Me.areceptor & "_DB.accdb'" & _
" WHERE Receptor = codigolocal()" & _
" UNION ALL SELECT *,codigolocal() as Emisor FROM ChatPitaya IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & codigoLocal() & "_DB.accdb'" & _
" WHERE Receptor = " & Me.areceptor & ")x"
DoCmd.SetWarnings True

'usar tabla temporal como de formualrio
Me.RecordSource = "SELECT Fecha," & _
" 'C:\Users\' & NombreSistema() & '\Google Drive BP\Base de Datos Pitaya\Sys Resources\ChatPitaya\msg' & IIf(Emisor = codigolocal(), 1, 2) & '.png' AS Fondo," & _
" IIf(Emisor=codigolocal(),Mensaje) AS Lado2," & _
" IIf(Emisor=" & Me.areceptor & ",Mensaje) AS Lado1," & _
" IIf(Emisor=codigolocal(),Fecha) AS Fecha2," & _
" IIf(Emisor=" & Me.areceptor & ",Fecha) AS Fecha1" & _
" FROM TempMensajes" & _
" ORDER BY Fecha"

cantactual = Me.Form.RecordsetClone.RecordCount
If cantactual > 3 Then
    DoCmd.GoToRecord , , acGoTo, cantactual - 3 'cantmen
End If

Me.Fecha1.Visible = True
Me.Fecha2.Visible = True
Me.Lado1.Visible = True
Me.Lado2.Visible = True

End Sub

Private Sub Form_Close()
Me.RecordSource = ""
Me.areceptor.RowSource = ""
DoCmd.RunSQL "DROP TABLE temp_chatsistema"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.InsideHeight = 8000
Me.Mensaje.ForeColor = RGB(169, 169, 169)
Me.Mensaje = "Maximo 72 caracteres"

DoCmd.SetWarnings False
DoCmd.RunSQL "SELECT x.Nombre & ' ' & x.Ciudad AS Name, x.CodSistema INTO temp_chatsistema" & _
" FROM (" & temporalarrayunion("DatosSistema") & ")x" & _
" WHERE Not x.CodSistema=codigolocal()"
DoCmd.SetWarnings True

Me.areceptor.RowSource = "SELECT * FROM temp_chatsistema"
Me.areceptor.RowSourceType = "Table/Query"
Me.areceptor.BoundColumn = 2

Call actualizarlista

End Sub

Private Sub areceptor_Change()
Call actualizarlista
End Sub

Private Sub enviar_Click()

If Not Me.Mensaje = "Maximo 72 caracteres" And Not Me.Mensaje = "" Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO ChatPitaya ( Mensaje, Receptor, Fecha) IN 'C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & codigoLocal() & "_DB.accdb'" & _
    " VALUES ( '" & Me.Mensaje & "', " & Me.areceptor & ", #" & Now & "#)"
    DoCmd.SetWarnings True
    
    Me.Mensaje.ForeColor = RGB(169, 169, 169)
    Me.Mensaje = "Maximo 72 caracteres"
    Call actualizarlista
End If

End Sub

Private Sub Form_Timer()

Dim cantguardado As Integer
cantguardado = DCount("[Emisor]", "TempMensajes", "[Emisor] = " & Me.areceptor & " AND DateValue([Fecha]) = Date()")
If Not cantguardado = mensajesdialocal(Me.areceptor) Then
    MsgBox "nuevo mensaje"
    Call actualizarlista
End If

End Sub

Private Sub Mensaje_Click()
Me.Mensaje = ""
Me.Mensaje.ForeColor = RGB(0, 0, 0)
End Sub


Private Sub Mensaje_Exit(Cancel As Integer)
If Me.Mensaje = "Maximo 72 caracteres" Or Me.Mensaje = "" Then
    Me.Mensaje.ForeColor = RGB(169, 169, 169)
    Me.Mensaje = "Maximo 72 caracteres"
End If
End Sub
