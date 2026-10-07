' ==========================================================
' Modulo  : Form_Cierre por Turno 3
' Tipo    : 100
' Lineas  : 395
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database
Public Sub actualizariconoscierre()

Dim contad As Integer
Dim tex As String
Dim resulta As String
Me.iconopos.Visible = False
Me.iconotrans.Visible = False
Me.iconopya.Visible = False
Me.iconoefectivo.Visible = False
tex = ""
resulta = "Diferencia en "
contad = 1


If Me.varpos <> 0 Then
    Me.iconopos.Visible = True
    tex = tex & contad & ". " & "Diferencia en POS Banco, requiere revision" & Chr(13) & Chr(10)
    resulta = resulta & IIf(contad <> 1, " / ", "") & "POS " & absolutoabsoluto(Me.varpos) & "C$"
    contad = contad + 1
End If
If Me.vartrans <> 0 Then
    Me.iconotrans.Visible = True
    tex = tex & contad & ". " & "Diferencia en Transferencias, requiere revision" & Chr(13) & Chr(10)
    resulta = resulta & IIf(contad <> 1, " / ", "") & "TRANSFERENCIA " & absolutoabsoluto(Me.vartrans) & "C$"
    contad = contad + 1
End If
If Me.varpya <> 0 Then
    Me.iconopya.Visible = True
    tex = tex & contad & ". " & "Diferencia en PedidosYa, requiere revision" & Chr(13) & Chr(10)
    resulta = resulta & IIf(contad <> 1, " / ", "") & "PEDIDOS YA " & absolutoabsoluto(Me.varpya) & "C$"
    contad = contad + 1
End If
If Me.varefectivo <> 0 Then
    Me.iconoefectivo.Visible = True
    tex = tex & contad & ". " & IIf(Me.totalentregado - Me.totalentregar > 0, "SOBRANTE DE EFECTIVO ", "FALTANTE DE EFECTIVO") & " de " & Me.faltaefectivo & "C$" & Chr(13) & Chr(10)
    contad = contad + 1
End If
Me.advertencias = tex
Me.resultadofinal = IIf(resulta = "Diferencia en ", "", resulta & ", requiere revision")

Me.comprascaja.Requery
Me.poscalculado.Requery
Me.transcalculado.Requery
Me.pyacalculado.Requery
Me.efectivocalculado.Requery
Me.Requery
End Sub


Private Sub Comando1192_Click()
On Error GoTo Nulo
Dim nuevo As Double
nuevo = InputBox("Ingrese nuevo valor de resumen detallado del POS de banco:", "Cambiar Monto", 0)
Me.posguardado = nuevo
Call actualizariconoscierre
Exit Sub
Nulo:

End Sub

Private Sub Comando1218_Click()
On Error GoTo Nulo
Dim nuevo As Double
nuevo = InputBox("Ingrese nuevo valor del monto total de pedidos aprobados sacado del equipo de PedidosYa:", "Cambiar Monto", 0)
Me.pyaguardado = nuevo
Call actualizariconoscierre
Exit Sub
Nulo:

End Sub

Private Sub Comando1220_Click()
On Error GoTo Nulo
Dim nuevo As Double
nuevo = InputBox("Ingrese nuevo valor de la suma de transferencias ralizadas en el dia:", "Cambiar Monto", 0)
Me.transguardado = nuevo
Call actualizariconoscierre
Exit Sub
Nulo:

End Sub
Private Sub Comando1207_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Date & "# AND [POS]<>0 AND [Delivery]=0 AND [Transferencia]=0 AND [Hora]>#" & Me.cinicio & "# AND [Hora]<#" & Me.cfinal & "#"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & " POS"
[Forms]![HistorialVentasFiltro].tipofiltro = "POS"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False
End Sub


Private Sub Comando1225_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Date & "# AND [POS]=0 AND [Comision]=0 AND [Transferencia]=0 AND [Hora]>#" & Me.cinicio & "# AND [Hora]<#" & Me.cfinal & "#"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & " EFECTIVO"
[Forms]![HistorialVentasFiltro].tipofiltro = "EFECTIVO"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False

End Sub

Private Sub Comando1227_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Date & "# AND [POS]<>0 AND ([Delivery]=5 Or [Delivery]=6 Or [Delivery]=8) AND [Hora]>#" & Me.cinicio & "# AND [Hora]<#" & Me.cfinal & "#"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & " PEDIDOS YA"
[Forms]![HistorialVentasFiltro].tipofiltro = "PEDIDOSYA"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False

End Sub

Private Sub Comando1229_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal, , "[Fecha]=#" & Date & "# AND [POS]<>0 AND [Delivery]=0 AND [Transferencia]<>0 AND [Hora]>#" & Me.cinicio & "# AND [Hora]<#" & Me.cfinal & "#"
[Forms]![HistorialVentasFiltro].tituloprincipal.Caption = "HISTORIAL DE PEDIDOS DIARIOS " & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & " TRANSFERENCIA"
[Forms]![HistorialVentasFiltro].tipofiltro = "TRANSFERENCIA"
[Forms]![HistorialVentasFiltro].semanaactual.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta216.Visible = False
[Forms]![HistorialVentasFiltro].asemana.Visible = False
[Forms]![HistorialVentasFiltro].Etiqueta346.Visible = False
[Forms]![HistorialVentasFiltro].Comando102.Visible = False
[Forms]![HistorialVentasFiltro].Comando133.Visible = False

End Sub

Private Sub Comando1281_Click()
DoCmd.OpenForm "Ingreso de Compras"
[Forms]![Ingreso de Compras]![operarioorigen] = Me.ccodigocajero
End Sub

Private Sub Comando1290_Click()
DoCmd.OpenForm "Contador Caja"
[Forms]![Contador Caja]![titulo1] = "NUEVO CONTEO DE CAJA"
[Forms]![Contador Caja]![titulo2] = "CIERRE " & Me.ccodigocierre

End Sub

Private Sub Comando426_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO CierreDiario(HoraInicial, HoraFinal, Fecha, CodOperario," & _
" MFCor, MFDol, Faltante, TotalHugo," & _
" TotalPedidosYa, TotalTransferencia, TotalPOS, Observaciones)" & _
" values (#" & Me.cinicio & "#, #" & Me.cfinal + 1 / 24 / 60 & "#, #" & Me.cfecha & "#, " & Me.ccodigocajero & "," & _
" " & Me.totalcordobas & ", " & Me.totaldolares & ", " & Me.totalentregado - Me.totalentregar & ", 0," & _
" " & Me.pyaguardado & ", " & Me.transguardado & ", " & Me.posguardado & ", '" & Nz(Me.Observaciones, " ") & "')"
DoCmd.SetWarnings True

'Historial de ventas POS
DoCmd.OpenReport "HistorialVentasxPedidoDia", acViewReport, , "[Pago]='POS' AND [Fecha]=#" & Date & "#"
DoCmd.SelectObject acReport, "HistorialVentasxPedidoDia"
DoCmd.PrintOut acSelection, 1, 1
DoCmd.Close acReport, "HistorialVentasxPedidoDia"

DoCmd.OpenReport "Cierre por Turno 2", acViewNormal
DoCmd.OpenReport "ResumenVentasMostradorDia", acViewNormal
DoCmd.Close acForm, "Cierre por Turno 2"
DoCmd.Close acForm, "Contador Caja 2"


MsgBox "Documentos impresos, se procedera a subir infromacion a la web, veriifcar conexion a internet"

'Ingreso de compras automatico
Dim fechaactual As Date
fechaactual = Date
Call IngresoComprasAutomatico(fechaactual)


''''''ENVIO REPORTES A TELEGRAM'''''

'Dim saltolinea As String
'saltolinea = "----------------------------------------------------------"

'''''''''''''''''''''''''''''''''''''''''''''''''''
'Reporte de cierre a telegram SIEMPRE a aprtir de 9 julio se dejara de enviar para agilizar

'Dim mensax1 As String

'mensax1 = saltolinea & vbCrLf & "CIERRE DE CAJA " & Me.ccodigocierre & vbCrLf & saltolinea & vbCrLf & _
          "Responsable de Cierre: " & Me.nombrecajero & vbCrLf & saltolinea
'Call EnviarMensajeTelegram(mensax1, grupotgerencia())

''''''''''''''''''''''''''''''''''''''''''''
'Dim mensax2 As String
'Dim Transferencia, pedidosya, totales, direfectivo, dircuenta, Otros As Double
'Dim reporte As String

'dircuenta = Me.poscalculado
'Transferencia = Me.transcalculado
'pedidosya = Me.pyacalculado
'direfectivo = Me.efectivocalculado
'totales = Me.TotalCalculado
'Otros = Round(totales - dircuenta - direfectivo - pedidosya - Transferencia, 0)

'reporte = "+ Ventas Totales: " & totales & vbCrLf & _
    saltolinea & vbCrLf & _
    "   Ventas por PedidosYa: " & pedidosya & vbCrLf & _
    "   Venta en efectivo: " & direfectivo & vbCrLf & _
    "   Venta con tarjeta: " & dircuenta & vbCrLf & _
    "   Venta con transferencia: " & Transferencia & vbCrLf & _
    "   Otros: " & Otros & vbCrLf

'mensax2 = "" '"+ Caja Inicial: " & cajainicial(Date) & vbCrLf & _
          'reporte & vbCrLf & _
          '"- Aligeramientos: " & Me.aligeramiento & vbCrLf & saltolinea & _
          '"Total a Entregar: " & Me.totalentregar & vbCrLf & saltolinea
'Call EnviarMensajeTelegram(mensax2, grupotgerencia())

'''''''''''''''''''''''''''''''''''''''''''''''''''
'Call EnviarMensajeTelegram("Resultado por Transferencia, Pediosya o POS: " & [Forms]![Cierre por Turno]![titulosobrafalta] & " de " & [Forms]![Cierre por Turno]![cantidadsobrafalta], grupotoperaciones())
'''''''''''''''''''''''''''''''''''''''''''''''''''

'Dim reporte8 As String

'reporte8 = "+ Efectivo Entregado: " & Me.totalentregado & vbCrLf & saltolinea & vbCrLf & _
    "   Total en Cordobas: " & Me.totalcordobas & vbCrLf & _
    "   Total en Dolares : " & Me.totaldolares & vbCrLf & _
    "+ Compras de Caja: " & Me.comprascaja & vbCrLf & _
    "+ Ventas Fisicas POS: " & Me.posguardado & vbCrLf & _
    "+ Ventas Fisicas Transferencia: " & Me.transguardado & vbCrLf & _
    "+ Ventas Fisicas PedidosYa: " & Me.pyaguardado & vbCrLf & _
    "Total Entregado: " & Me.montoperiodo & vbCrLf & saltolinea
'Call EnviarMensajeTelegram(reporte8, grupotgerencia())

'''''''''''''''''''''''''''''''''''''''''''''''''''
'Call EnviarMensajeTelegram("Resultado de Efectivo: " & Me.tiposobrafalta & " de " & Me.sobrafalta, grupotoperaciones())
'Call EnviarMensajeTelegram("Resultado de Cierre: " & Me.Etiqueta1075 & " " & Me.faltaefectivo, grupotgerencia())

'Dim horall As Date
'horall = Time()

'If Hour(horall) >= 17 Then 'cierres despues de las 5


   'Reporte de cierre a telegram ULTIMO
    ''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''VENTAS
    'Dim mediano As Double
    'Dim grande As Double
    'Dim reporte10 As String
    'mediano = PorcentajeMedidasBatidos("Mediano", Date) * 100
    'grande = PorcentajeMedidasBatidos("Gigantona", Date) * 100

    'reporte10 = saltolinea & vbCrLf & "DETALLE DE VENTAS" & vbCrLf & _
        "  Mediano: " & Round(mediano) & "%" & vbCrLf & _
        "  Grande: " & Round(grande) & "%" & vbCrLf & saltolinea & vbCrLf

    'Dim especi, premiu, clasic, tarje, bowl, pitayasto, salud, adiciona, waffle, limon, parfa, cafechoco, protein As Double
    'Dim reporte2 As String
    'especi = CantidadVentasGrupoDia(Date, 1)
    'premiu = CantidadVentasGrupoDia(Date, 2)
    'clasic = CantidadVentasGrupoDia(Date, 3)
    'tarje = CantidadVentasGrupoDia(Date, 5)
    'bowl = CantidadVentasGrupoDia(Date, 6)
    'pitayasto = CantidadVentasGrupoDia(Date, 7)
    'salud = CantidadVentasGrupoDia(Date, 8)
    'protein = CantidadVentasGrupoDia(Date, 24)
    'adiciona = CantidadVentasGrupoDia(Date, 11)
    'waffle = CantidadVentasGrupoDia(Date, 14)
    'limon = CantidadVentasGrupoDia(Date, 16)
    'parfa = CantidadVentasGrupoDia(Date, 17)
    'cafechoco = CantidadVentasGrupoDia(Date, 18) + CantidadVentasGrupoDia(Date, 19)

    'reporte2 = "  Batidos Premium:  " & premiu & vbCrLf & _
        "  Batidos Saludables: " & salud & vbCrLf & _
        "  Batidos Proteina: " & protein & vbCrLf & _
        "  Batidos Especiales:  " & especi & vbCrLf & _
        "  Batidos Clasicos: " & clasic & vbCrLf & _
        "  Tarjetas:  " & tarje & vbCrLf & _
        "  Smoothie Bowl:  " & bowl & vbCrLf & _
        "  Parfait:  " & parfa & vbCrLf & _
        "  Limonadas:  " & limon & vbCrLf & _
        "  Waffle:  " & waffle & vbCrLf & _
        "  Chocolates y Cafe:  " & cafechoco & vbCrLf & _
        "  Adicionales: " & adiciona & vbCrLf & _
        "  Pitaya Store:  " & pitayasto & vbCrLf & saltolinea & vbCrLf

    'Dim cantipromocorte As Double
    'Dim cantipromodevol As Double
    'Dim reporte6 As String
    'cantipromocorte = CantidadProductosxPromocionDia(Date, 6)
    'cantipromodevol = CantidadProductosxPromocionDia(Date, 39)
    'reporte6 = "  Cortesias: " & cantipromocorte & vbCrLf & _
    "  Devoluciones: " & cantipromodevol & vbCrLf & saltolinea & vbCrLf
    
    'Dim tickpro, cantiva As Double
    'Dim reporte3 As String
    'cantiva = CantidadPedidosValidosDia(Date)
    'If cantiva = 0 Then tickpro = 0 Else tickpro = Round(totales / cantiva, 0)
    'reporte3 = "Ticket Promedio:" & tickpro & vbCrLf & saltolinea & vbCrLf
    
    'Call EnviarMensajeTelegram(reporte10 & reporte2 & reporte6 & reporte3, grupotoperaciones())
    'Call EnviarMensajeTelegram(reporte10 & reporte2 & reporte6 & reporte3, grupotgerencia())
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Dim mermacoti As String
    'Dim mermaingre As String
    'Dim reporte5 As String

    'mermacoti = ResumenMermasCotiDia(Date)
    'mermaingre = ResumenMermasIngreDia(Date)
    'reporte5 = saltolinea & vbCrLf & "DETALLE DE MERMAS" & vbCrLf & saltolinea & _
               "Mermas del Dia: " & vbCrLf & mermacoti & vbCrLf & mermaingre
    'Call EnviarMensajeTelegram(reporte5, grupotgerencia())

    'Me.Requery

    '''''''''''''''''''''''''''''''''''''''''''''''Imprimir graficas
    'Dim acuma As Integer
    'Dim acumm As Integer
    'acuma = Year(Date)
    'acumm = Month(Date)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''acumulado mes
    'On Error Resume Next
    'Dim Direccionano, Direccionmes As String
    'Dim archivo As String

    'Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & acuma
    'Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & acuma & "\" & acumm
    'If Dir(Direccionano, vbDirectory) = "" Then
    '    MkDir Direccionano
    'End If
    'If Dir(Direccionmes, vbDirectory) = "" Then
    '    MkDir Direccionmes
    'End If
    
    'archivo = Direccionmes & "\P" & codigolocal() & " - Acumulado Meta Mensual.jpg"
    'DoCmd.OpenForm "Dashboard_Ventas_DiaMesAcumulado", acFormPivotChart, , "[ames]=" & acumm & " AND [aano]=" & acuma & " AND [Dates]<=#" & Date & "#"
    '[Forms]![Dashboard_Ventas_DiaMesAcumulado].Form.Requery
    '[Forms]![Dashboard_Ventas_DiaMesAcumulado].ChartSpace.exportpicture archivo
    'DoCmd.Close acForm, "Dashboard_Ventas_DiaMesAcumulado"
    
    'Me.Requery
    'Call EnviarImagenTelegram(Direccionmes & "\", "P" & codigolocal() & " - Acumulado Meta Mensual.jpg", "Acumulado Meta Mensual " & nombrelocal(), grupotoperaciones())
    'Call EnviarImagenTelegram(Direccionmes & "\", "P" & codigolocal() & " - Acumulado Meta Mensual.jpg", "Acumulado Meta Mensual " & nombrelocal(), grupotgerencia())
    'On Error GoTo 0

    '''''''''''''''''''''''''''''''''''''''''''''''''''diario con meta
    'On Error Resume Next
    'Dim Direccionanos, Direccionmess As String
    'Dim archivo2 As String

    'Direccionanos = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & acuma
    'Direccionmess = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & acuma & "\" & acumm
    'If Dir(Direccionanos, vbDirectory) = "" Then
    '    MkDir Direccionanos
    'End If
    'If Dir(Direccionmess, vbDirectory) = "" Then
    '    MkDir Direccionmess
    'End If

    'archivo2 = Direccionmess & "\P" & codigolocal() & " - Dia a Dia Mensual.jpg"
    'DoCmd.OpenForm "Dashboard_Ventas_DiaMes", acFormPivotChart, , "[ames]=" & acumm & " AND [aano]=" & acuma & " AND [Dates]<=#" & Date & "#"
    '[Forms]![Dashboard_Ventas_DiaMes].Form.Requery
    '[Forms]![Dashboard_Ventas_DiaMes].ChartSpace.exportpicture archivo2
    'DoCmd.Close acForm, "Dashboard_Ventas_DiaMes"

    'Me.Requery
    'Call EnviarImagenTelegram(Direccionmess & "\", "P" & codigolocal() & " - Dia a Dia Mensual.jpg", "Dia a Dia Mensual " & nombrelocal(), grupotoperaciones())
    'Call EnviarImagenTelegram(Direccionmess & "\", "P" & codigolocal() & " - Dia a Dia Mensual.jpg", "Dia a Dia Mensual " & nombrelocal(), grupotgerencia())
    'On Error GoTo 0
'End If

If EsSistemaDeTienda() Then ' solo si no es sistema raiz, es sistema tienda
    Call SincronizarDatosClientesLocales
    Call ActualizarArchivoCSVVentasFecha(Date)
    Call ActualizarArchivoCSVVentasFecha(Date - 1)
    Call ActualizarArchivoCSVMembresias
    Call SyncKardexTiendaCierre30Dias
    Call SyncCierreDepositosTienda30Dias
End If

DoCmd.Close acForm, "Cierre por Turno 3"

'Application.FollowHyperlink "https://erp.batidospitaya.com"
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Call actualizariconoscierre
End Sub

