' ==========================================================
' Modulo  : Form_Corroborar Ingresos Pitaya
' Tipo    : 100
' Lineas  : 107
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando147_Click()
If Comando147.Caption = "QUITAR CEROS" Then
    If codigoLocal() <> 0 Then 'Sistema local
        Me.Filter = "[Global]=0 AND ([Compras] <> 0 OR [Ingresos] <> 0)"
    Else
        Me.Filter = "[Compras] <> 0 OR [Ingresos] <> 0"
    End If
    Me.FilterOn = True
    Me.Requery
    Me.Comando147.Caption = "MOSTRAR TODO"
Else
    If codigoLocal() <> 0 Then 'Sistema local
        Me.Filter = "[Global]=0"
    Else
        Me.Filter = ""
    End If
    Me.FilterOn = True
    Me.Requery
    Me.Comando147.Caption = "QUITAR CEROS"
End If
End Sub

Public Sub Comando170_Click()
Me.Requery
Me.Filter = "[Compras] <> 0 OR [Ingresos] <> 0 OR [PreIngre] <> 0"
Me.FilterOn = True
Me.Comando147.Caption = "MOSTRAR TODO"

'Dim printcart As String
'impresoracarta
'impresorapapeltermico
'impresorastickertermico
'printcart = DLookup("[impresoracarta]", "DatosSistema")

'Dim prtAvailPrinters As Printer

'For Each prtAvailPrinters In Application.Printers
'    If prtAvailPrinters.DeviceName = sPrinterName Then
'        Set Application.Printer = prtAvailPrinters
'        Exit For
'    End If
'Next prtAvailPrinters

Call configurarmargentamañoformulario("Corroborar Ingresos Pitaya")

Dim Direccion, Direccion2 As String
Dim archivo As String

If codigoLocal() = 0 Then
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Almacen\Reporte Semanal\" & Me.semanaac
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Almacen\Reporte Semanal"

Else
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.semanaac
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
End If

If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Control de Ingresos.pdf"
DoCmd.OutputTo acOutputForm, "Corroborar Ingresos Pitaya", acFormatPDF, archivo, False

DoCmd.Close acForm, "Corroborar Ingresos Pitaya"
End Sub

Private Sub Comando200_Click()


End Sub

Private Sub Comando221_Click()
DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").modosucursal
End Sub



Private Sub Comando65_Click()
Me.Requery
Me.Filter = "[Compras] <> 0 OR [Ingresos] <> 0 OR [PreIngre] <> 0"
Me.FilterOn = True
Me.Comando147.Caption = "MOSTRAR TODO"

End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaant = numerosemana(Date)
'Call importarpreingresossistema0
Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)
End Sub
