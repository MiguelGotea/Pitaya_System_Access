' ==========================================================
' Modulo  : Form_Estado de Resultados Local
' Tipo    : 100  |  Lineas: 119
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando161_Click()
Me.Requery
End Sub

Private Sub calculardatos_Click()
Me.mesanterior.Requery
Me.anoanterior.Requery

Me.ventasteoricas = IngresoMesTeorico(Me.mesac, Me.añoac)
Me.ventasejecutadas = FRVentaTotalMesLocal(Me.mesac, Me.añoac, 3)

Me.cierres = SumaCierresMesInterno(Me.mesac, Me.añoac)
Me.cajainicial = SumaCajaInicialMesInterno(Me.mesac, Me.añoac)
Me.comprascaja = SalidasDeCajaMesInterno(Me.mesac, Me.añoac)
Me.ventaspos = FRVentaTotalMesLocal(Me.mesac, Me.añoac, 1)

Me.costoventasteorico = FRCostoConsumoTeoricoMes(Me.mesac, Me.añoac)
Me.ING = FRCostoIngresosMes(Me.mesac, Me.añoac)
Me.MERMA = FRCostoMermasCotizacionMes(Me.mesac, Me.añoac) + FRCostoMermasIngredienteMes(Me.mesac, Me.añoac)

'==========================Inventarios Finales de mes============================
Dim feim As Date
Dim feim2 As Date
Dim fefm As Date
Dim fefm2 As Date
Dim usem, usem2 As Integer

feim = ultimodomingomes(Me.mesanterior, Me.anoanterior) + 1
fefm = DateSerial(Me.anoanterior, Me.mesanterior, diasmes(DateSerial(Me.anoanterior, Me.mesanterior, 1)))
usem = numerosemana(feim)

If feim - 1 = fefm Then
    Me.intervalo1 = "Ultimo domingo coincide con fin de mes " & fefm
    Me.CIIIFUD = FRCostoIFIngrediente(usem - 1) + FRCostoIFCotizacion(usem - 1)
    Me.CIIIUD = 0
    Me.CIICUD = 0
    Me.CIIMUD = 0
Else
    Me.intervalo1 = "Desde " & feim & " hasta " & fefm
    Me.CIIIFUD = FRCostoIFIngrediente(usem - 1) + FRCostoIFCotizacion(usem - 1)
    Me.CIIIUD = FRCostoIngresoIntervalo(feim, fefm)
    Me.CIICUD = FRCostoConsumoIntervalo(feim, fefm)
    Me.CIIMUD = FRCostoMermaIngredienteIntervalo(feim, fefm) + FRCostoMermaCotizacionIntervalo(feim, fefm)
End If


feim2 = ultimodomingomes(Me.mesac, Me.añoac) + 1
fefm2 = DateSerial(Me.añoac, Me.mesac, diasmes(DateSerial(Me.añoac, Me.mesac, 1)))
usem2 = numerosemana(feim2)

If feim2 - 1 = fefm2 Then
    Me.intervalo2 = "Ultimo domingo coincide con fin de mes " & fefm2
    Me.CIFIFUD = FRCostoIFIngrediente(usem2 - 1) + FRCostoIFCotizacion(usem2 - 1)
    Me.CIFIUD = 0
    Me.CIFCUD = 0
    Me.CIFMUD = 0
Else
    Me.intervalo2 = "Desde " & feim2 & " hasta " & fefm2
    Me.CIFIFUD = FRCostoIFIngrediente(usem2 - 1) + FRCostoIFCotizacion(usem2 - 1)
    Me.CIFIUD = FRCostoIngresoIntervalo(feim2, fefm2)
    Me.CIFCUD = FRCostoConsumoIntervalo(feim2, fefm2)
    Me.CIFMUD = FRCostoMermaIngredienteIntervalo(feim2, fefm2) + FRCostoMermaCotizacionIntervalo(feim2, fefm2)
End If
'==================================================================================

Me.cgene = ComprasGrupoMesControl(Me.mesac, Me.añoac, "FIJOS", "Generales")
Me.crepo = ComprasGrupoMesControl(Me.mesac, Me.añoac, "FIJOS", "Reposicion")
Me.gprom = Me.ventasteoricas - Me.ventasejecutadas
Me.gmant = ComprasGrupoMesNoControl(Me.mesac, Me.añoac, "FIJOS", "Mantenimiento")
Me.gpers = ComprasGrupoMesNoControl(Me.mesac, Me.añoac, "FIJOS", "Personal")
Me.grent = ComprasGrupoMesNoControl(Me.mesac, Me.añoac, "FIJOS", "Renta")
Me.gserv = ComprasGrupoMesNoControl(Me.mesac, Me.añoac, "FIJOS", "Servicios")
Me.gotro = ComprasGrupoMesNoControl(Me.mesac, Me.añoac, "FIJOS", "Otros Fijos")
End Sub

Private Sub cpersonal_Click()
DoCmd.OpenForm "Busqueda Compras", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Personal'"
End Sub
Private Sub cgenerales_Click()
DoCmd.OpenForm "Historial Ingresos", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Generales'"
End Sub
Private Sub creposicion_Click()
DoCmd.OpenForm "Historial Ingresos", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Reposicion'"
End Sub
Private Sub crenta_Click()
DoCmd.OpenForm "Busqueda Compras", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Renta'"
End Sub
Private Sub cmantenimiento_Click()
DoCmd.OpenForm "Busqueda Compras", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Mantenimiento'"
End Sub
Private Sub cservicios_Click()
DoCmd.OpenForm "Busqueda Compras", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Servicios'"
End Sub
Private Sub cotros_Click()
DoCmd.OpenForm "Busqueda Compras", , , "[Mes]=" & Me.mesac & " AND [Año]=" & Me.añoac & " AND [TIPO1]='FIJOS' AND [TIPO2]= 'Otros Fijos'"
End Sub

Private Sub Comando574_Click()
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
archivo = Direccionmes & "\" & Me.mesac & "_" & Me.añoac & " - Estado de Resultado P" & codigoLocal() & ".pdf"
DoCmd.OutputTo acOutputForm, "Estado de Resultados Local", acFormatPDF, archivo, True

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
