' ==========================================================
' Modulo  : Form_Flujo Caja Mensual Global Costos
' Tipo    : 100
' Lineas  : 48
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
' ==========================================================
Option Compare Database
Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 10500
End Sub

Private Sub Comando562_Click()
Dim tip1, tip2 As String
Dim Mes, ano As Integer
tip1 = Me.TIPO1
tip2 = Me.TIPO2
Mes = mesac
ano = añoac

DoCmd.OpenForm "HistorialComprasxTipoFlujoCaja", acNormal
Forms("HistorialComprasxTipoFlujoCaja").Filter = "[TIPO1]= '" & tip1 & "' and [TIPO2]= '" & tip2 & "' and [ames]= " & Mes & " and [aano]= " & ano
Forms("HistorialComprasxTipoFlujoCaja").FilterOn = True

End Sub

Private Sub Comando574_Click()
Call Comando65_Click
Dim Direccionano, Direccionmes As String
Dim archivo As String
Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.añoac
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.añoac & "\" & Me.mesac
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If
archivo = Direccionmes & "\" & Me.mesac & "_" & Me.añoac & " - Costos Acumulados Totales.pdf"
DoCmd.OutputTo acOutputForm, "Flujo Caja Mensual Global Costos", acFormatPDF, archivo, True

End Sub

Private Sub Comando65_Click()
Me.Requery

Me.sistema0 = ComprasGlobalSistema0(Me.mesac, Me.añoac)
Me.porcionglobal = ComprasGlobalMesAlmacenGlobal(Me.mesac, Me.añoac)
Me.egresosactual = ComprasGlobalMes(Me.mesac, Me.añoac) '+ Me.sistema0 '- Me.porcionglobal pendiente verificar compra de porcoines


End Sub

