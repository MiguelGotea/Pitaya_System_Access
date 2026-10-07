' ==========================================================
' Modulo  : Form_EleccionSucursalVigente
' Tipo    : 100
' Lineas  : 41
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database


Private Sub Comando48_Click()
Dim loca As Integer
Dim indi As String
If IsNull(Me.alocal) Or Me.alocal = "" Then
    MsgBox "ingresar sucursal correctamente"
    Exit Sub
End If

loca = Me.alocal
Me.alocal.RowSource = ""
Select Case CurrentProject.Name
    Case "Pitaya_System.accdb", "Pitaya_System - Copy.accdb", "Pitaya_System - Copia.accdb" 'archivo base de sistema
        Call actualizartablas(Me.tipocarga, loca)
    Case Else ' archivo de local especifico
        MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
End Select

If CurrentProject.AllForms("Menu Gestion").IsLoaded Then
    [Forms]![Menu Gestion].[Sucursal] = ciudadsistema()
    [Forms]![Menu Gestion].[actualizacion] = fechaultimasubidaarchivodrive("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & codigoLocal() & "_DB.accdb")
End If
If CurrentProject.AllForms("Introduccion").IsLoaded Then
    [Forms]![Introduccion].[aloca] = loca
End If

MsgBox "Sistema Cargado Correctamente"
DoCmd.Close acForm, "EleccionSucursalVigente"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "ELEGIR SUCURSAL"
Me.ShortcutMenu = False
Call importartablaespecifica("Main_DB", "StatusSucursales", "StatusSucursales", 1)
Call importartablaespecifica("Main_DB", "TablasModulos", "TablasModulos", 1)
End Sub



