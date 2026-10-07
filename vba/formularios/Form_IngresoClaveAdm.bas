' ==========================================================
' Modulo  : Form_IngresoClaveAdm
' Tipo    : 100  |  Lineas: 123
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database

Private Sub botonbloquear_Click()

Select Case Me.clavedesbloquear
    'Funciones casos de abrir modulo gestion
    Case "nihonkoperaciones"
        DoCmd.OpenForm "Menu Gestion"
        [Forms]![Menu Gestion]![main].Pages("Ventas").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Costos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Consumos").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Registros").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Balances").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Personal").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Descargar").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Administracion").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Resumen").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Tools").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Sistema").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Global").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Rev Mes").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Plan").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Central").Visible = False
        
        DoCmd.Close acForm, "Introduccion"
    Case "nihonkcds"
        Select Case CurrentProject.Name
            Case "Pitaya_System.accdb" 'archivo base de sistema
                Call actualizartablas(1, 0)
            Case Else ' archivo de local especifico
                MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
        End Select
        DoCmd.OpenForm "Menu Gestion"
        [Forms]![Menu Gestion]![main].Pages("Ventas").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Costos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Consumos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Registros").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Balances").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Personal").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Descargar").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Administracion").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Resumen").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Tools").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Sistema").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Global").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Rev Mes").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Plan").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Central").Visible = True
        [Forms]![Menu Gestion]![Comando3849].Visible = False
        DoCmd.Close acForm, "Introduccion"
    Case "nihonksistema"
        DoCmd.OpenForm "Menu Gestion"
        [Forms]![Menu Gestion]![main].Pages("Ventas").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Costos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Consumos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Registros").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Balances").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Personal").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Descargar").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Administracion").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Resumen").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Tools").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Sistema").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Global").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Rev Mes").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Plan").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Central").Visible = False
        DoCmd.Close acForm, "Introduccion"
    Case "nihonkcontabilidad"
        DoCmd.OpenForm "Menu Gestion"
        [Forms]![Menu Gestion]![main].Pages("Ventas").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Costos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Consumos").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Registros").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Balances").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Personal").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Descargar").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Administracion").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Resumen").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Tools").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Sistema").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Global").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Rev Mes").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Plan").Visible = False
        [Forms]![Menu Gestion]![main].Pages("Central").Visible = False
        DoCmd.Close acForm, "Introduccion"
    Case "nihonkmiguel"
        DoCmd.OpenForm "Menu Gestion"
        [Forms]![Menu Gestion]![main].Pages("Ventas").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Costos").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Consumos").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Registros").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Balances").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Personal").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Descargar").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Administracion").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Resumen").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Tools").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Sistema").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Global").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Rev Mes").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Plan").Visible = True
        [Forms]![Menu Gestion]![main].Pages("Central").Visible = True
        DoCmd.Close acForm, "Introduccion"
    Case Else
        MsgBox "Clave Erronea"
        Me.clavedesbloquear = ""
        Me.clavedesbloquear.SetFocus
End Select

DoCmd.Close acForm, "IngresoClaveAdm"

End Sub

Private Sub Comando1253_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

