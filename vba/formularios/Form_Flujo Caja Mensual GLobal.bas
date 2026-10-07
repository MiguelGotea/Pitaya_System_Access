' ==========================================================
' Modulo  : Form_Flujo Caja Mensual GLobal
' Tipo    : 100
' Lineas  : 176
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:18
' ==========================================================
Option Compare Database
Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 10500
End Sub
Private Sub Comando139_Click()
If mesac.Value = 1 Then
    mesac.Value = 12
    añoac.Value = añoac.Value - 1
    Me.ingresosactual = 0
    Me.egresosactual = 0
Else
    mesac.Value = mesac.Value - 1
    Me.ingresosactual = 0
    Me.egresosactual = 0
End If

End Sub

Private Sub Comando142_Click()
If mesac.Value = 12 Then
    mesac.Value = 1
    añoac.Value = añoac.Value + 1
    Me.ingresosactual = 0
    Me.egresosactual = 0
Else
    mesac.Value = mesac.Value + 1
    Me.ingresosactual = 0
    Me.egresosactual = 0
End If
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
Me.mesac.height = 0
Me.añoac.height = 0
Me.Comando574.height = 0
Me.nombreingreso1.height = 4760 '280cu

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
archivo = Direccionmes & "\" & Me.mesac & "_" & Me.añoac & " - Flujo de Caja.pdf"
DoCmd.OutputTo acOutputForm, "Flujo Caja Mensual Global", acFormatPDF, archivo, True

End Sub

Private Sub Comando65_Click()
Me.Requery

Me.cierres = SumaCierresMesGlobal(Me.mesac, Me.añoac)
Me.cajainicial = SumaCajaInicialMes(Me.mesac, Me.añoac)
Me.compras = SalidasDeCajaMes(Me.mesac, Me.añoac)

Me.sistema0 = ComprasGlobalSistema0(Me.mesac, Me.añoac)
Me.porcionglobal = ComprasGlobalMesAlmacenGlobal(Me.mesac, Me.añoac)
Me.egresosactual = ComprasGlobalMes(Me.mesac, Me.añoac) '+ Me.sistema0 '- Me.porcionglobal pendiente verificar compra de porcoines

On Error Resume Next
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim montpos, montnopos, canti, sobrantefaltante As Double
Dim totalxventa2pos, totalxventa2nopos, totalxventa4pos, totalxventa4nopos As Double

miSQL = "SELECT Grupos.prioridad, Grupos.CodGrupo, Grupos.Tipo, Grupos.NombreGrupo FROM Grupos ORDER BY Grupos.prioridad"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

Me.nombreingreso1 = ""
Me.nombreingreso2 = ""

'PITAYA 2
Me.ingresoslocales2 = ""
Me.montoslocales2pos = ""
Me.montoslocales2nopos = ""
totalxventa2pos = 0
totalxventa2nopos = 0
Me.ingresoxventas2 = 0
'PITAYA 4
Me.ingresoslocales4 = ""
Me.montoslocales4pos = ""
Me.montoslocales4nopos = ""
totalxventa4pos = 0
totalxventa4nopos = 0
Me.ingresoxventas4 = 0

DoCmd.SetWarnings False
For I = 1 To Cant
    Me.nombreingreso1 = Me.nombreingreso1 & rst("Tipo") & Chr(13) & Chr(10)
    Me.nombreingreso2 = Me.nombreingreso2 & rst("NombreGrupo") & Chr(13) & Chr(10)

    'PITAYA 2

    canti = FRVentaTotalMesGrupo(Me.mesac, Me.añoac, rst("CodGrupo"), 2, 2)
    montpos = FRVentaTotalMesGrupo(Me.mesac, Me.añoac, rst("CodGrupo"), 1, 2)
    'montnopos = FRVentaxGrupoMes(Me.mesac, Me.añoac, rst("CodGrupo"), 2, 2)

    Me.ingresoslocales2 = Me.ingresoslocales2 & Format(canti, "#,#0#") & Chr(13) & Chr(10)
    Me.montoslocales2pos = Me.montoslocales2pos & Format(montpos, "#,#0.0#") & Chr(13) & Chr(10)
    Me.ingresoxventas2 = Me.ingresoxventas2 + montpos
    'Me.montoslocales2nopos = Me.montoslocales2nopos & Format(montnopos, "#,#0.0#") & Chr(13) & Chr(10)
    'totalxventa2pos = totalxventa2pos + montpos
    'totalxventa2nopos = totalxventa2nopos + montnopos

    'PITAYA 4
    canti = FRVentaTotalMesGrupo(Me.mesac, Me.añoac, rst("CodGrupo"), 2, 4)
    montpos = FRVentaTotalMesGrupo(Me.mesac, Me.añoac, rst("CodGrupo"), 1, 4)
    'montnopos = FRVentaxGrupoMes(Me.mesac, Me.añoac, rst("CodGrupo"), 2, 4)

    Me.ingresoslocales4 = Me.ingresoslocales4 & Format(canti, "#,#0#") & Chr(13) & Chr(10)
    Me.montoslocales4pos = Me.montoslocales4pos & Format(montpos, "#,#0.0#") & Chr(13) & Chr(10)
    Me.ingresoxventas4 = Me.ingresoxventas4 + montpos
    'Me.montoslocales4nopos = Me.montoslocales4nopos & Format(montnopos, "#,#0.0#") & Chr(13) & Chr(10)
    'totalxventa4pos = totalxventa4pos + montpos
    'totalxventa4nopos = totalxventa4nopos + montnopos

    rst.MoveNext
Next I
DoCmd.SetWarnings True
rst.Close

'PITAYA 2
'Me.ingresoxventas2 = totalxventa2pos + totalxventa2nopos
'PITAYA 4
'Me.ingresoxventas4 = totalxventa4pos + totalxventa4nopos

Me.ingresoxventas = Me.ingresoxventas2 + Me.ingresoxventas4 'puede ser FRVentasTotalGruposMes(Me.mesac, Me.añoac, 2) + FRVentasTotalGruposMes(Me.mesac, Me.añoac, 1)

Dim ventatarjeta, ventaefectivo, ventahugo, ventapedidosya, ventaotro As Double
ventatarjeta = FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 0, -1, 1, 2) + FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 0, -1, 1, 4)
ventaefectivo = FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 0, 0, 1, 2) + FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 0, 0, 1, 4)
ventahugo = FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 5, -1, 1, 2) + FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 5, -1, 1, 4)
ventapedidosya = FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 8, -1, 1, 2) + FRVentaTotalMesTipoDeliv(Me.mesac, Me.añoac, 8, -1, 1, 4)
ventaotro = Me.ingresoxventas - ventatarjeta - ventaefectivo - ventahugo - ventapedidosya

Me.ventaspos = ventatarjeta 'totalxventa2pos + totalxventa4pos 'puede ser FRVentasTotalGruposMes(Me.mesac, Me.añoac, 1)
Me.ventashugo = ventahugo
Me.ventaspedidosya = ventapedidosya
Me.ventasotro = ventaotro
Me.ingresosactual = Me.cierres - Me.cajainicial + Me.compras + Me.ventaspos + Me.ventashugo + Me.ventaspedidosya + Me.ventasotro

sobrantefaltante = Me.ingresosactual - Me.ingresoxventas
If sobrantefaltante > 0 Then
    Me.sobrafaltaetiqueta.Caption = "SOBRANTE"
    Me.sobrafalta = sobrantefaltante
Else
    Me.sobrafaltaetiqueta.Caption = "FALTANTE"
    Me.sobrafalta = (-1) * sobrantefaltante
End If

End Sub

