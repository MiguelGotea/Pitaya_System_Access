' ==========================================================
' Modulo  : Form_Contador Caja
' Tipo    : 100  |  Lineas: 149
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database

Private Sub Comando823_Click()
Me.Requery

Select Case Me.titulo1
    Case "CONTEO DE CAJA INICIAL"  ' aplica cuando se hace conteo de caja inicial
        
        If codigoLocal() = 23 Or codigoLocal() = 14 Or EsSistemaDeTienda() = False Then
            DoCmd.OpenForm "Inicial"
            [Forms]![Inicial]![Dinero].Value = Me.ssubtotal
            [Forms]![Inicial].Form.SetFocus
            DoCmd.Close acForm, "Inicial"
            DoCmd.OpenReport "Contador Caja", acViewNormal
            'Application.FollowHyperlink "https://erp.batidospitaya.com"
        
        Else
        
            If MontoCierre(Date - 1, "T") <> Me.ssubtotal Then ' cuando se copia centro de formacion no pide caja inicial ni ferias
                MsgBox "El dinero ingresado no coincide con el monto dejado en el cierre anterior, verifficar conteo nuevamente"
                Exit Sub
            Else
                DoCmd.OpenForm "Inicial"
                [Forms]![Inicial]![Dinero].Value = Me.ssubtotal
                
                Call actualizardatosclubmixedglobal
                'Call DescargarTablaCompleta("clientesclub", "clientesclubexterno", "sucursal <> " & codigolocal())
                'Call EnviarMensajeTelegram("Iniciando con caja inicial " & Me.ssubtotal, grupotgerencia())
                
                [Forms]![Inicial].Form.SetFocus
                DoCmd.Close acForm, "Inicial"
                DoCmd.OpenReport "Contador Caja", acViewNormal
                'Application.FollowHyperlink "https://erp.batidospitaya.com"
    
            End If
        End If
        
    Case "NUEVO CONTEO DE CAJA" ' segundo conteo de caja de cierre
        DoCmd.OpenReport "Contador Caja", acViewNormal
        [Forms]![Cierre por Turno 3]![totalcordobas] = Me.scor
        [Forms]![Cierre por Turno 3]![totaldolares] = Me.sdol
        Call Forms("Cierre por Turno 3").actualizariconoscierre
        [Forms]![Cierre por Turno 3].Form.SetFocus
    Case Else
End Select


DoCmd.Close acForm, "Contador Caja"
End Sub




Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

Private Sub c0_5_Exit(Cancel As Integer)
If IsNull(Me.[c0.5]) Then
    Me.[c0.5] = 0
End If
End Sub

Private Sub c1_Exit(Cancel As Integer)
If IsNull(Me.c1) Then
    Me.c1 = 0
End If
End Sub

Private Sub c5_Exit(Cancel As Integer)
If IsNull(Me.c5) Then
    Me.c5 = 0
End If
End Sub

Private Sub c10_Exit(Cancel As Integer)
If IsNull(Me.c10) Then
    Me.c10 = 0
End If
End Sub

Private Sub c20_Exit(Cancel As Integer)
If IsNull(Me.c20) Then
    Me.c20 = 0
End If
End Sub

Private Sub c50_Exit(Cancel As Integer)
If IsNull(Me.c50) Then
    Me.c50 = 0
End If
End Sub

Private Sub c100_Exit(Cancel As Integer)
If IsNull(Me.c100) Then
    Me.c100 = 0
End If
End Sub

Private Sub c200_Exit(Cancel As Integer)
If IsNull(Me.c200) Then
    Me.c200 = 0
End If
End Sub

Private Sub c500_Exit(Cancel As Integer)
If IsNull(Me.c500) Then
    Me.c500 = 0
End If
End Sub

Private Sub d1_Exit(Cancel As Integer)
If IsNull(Me.d1) Then
    Me.d1 = 0
End If
End Sub

Private Sub d5_Exit(Cancel As Integer)
If IsNull(Me.d5) Then
    Me.d5 = 0
End If
End Sub

Private Sub d10_Exit(Cancel As Integer)
If IsNull(Me.d10) Then
    Me.d10 = 0
End If
End Sub

Private Sub d20_Exit(Cancel As Integer)
If IsNull(Me.d20) Then
    Me.d20 = 0
End If
End Sub

Private Sub d50_Exit(Cancel As Integer)
If IsNull(Me.d50) Then
    Me.d50 = 0
End If
End Sub

Private Sub d100_Exit(Cancel As Integer)
If IsNull(Me.d100) Then
    Me.d100 = 0
End If
End Sub

