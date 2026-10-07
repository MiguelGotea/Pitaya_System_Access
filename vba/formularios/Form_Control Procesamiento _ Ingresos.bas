' ==========================================================
' Modulo  : Form_Control Procesamiento / Ingresos
' Tipo    : 100  |  Lineas: 65
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:11
' ==========================================================

Option Compare Database

Private Sub Comando154_Click()
If Me.Comando154.Caption = "QUITAR CEROS" Then
    Me.Filter = "[SI] <> 0 or [INGRESOS] <> 0 or [PROCESADO] <> 0 or [MERMAS] <> 0 or [SF] <> 0"
    Me.FilterOn = True
    Me.Requery
    Me.Comando154.Caption = "MOSTRAR TODO"
Else
    Me.Filter = ""
    Me.FilterOn = True
    Me.Requery
    Me.Comando154.Caption = "QUITAR CEROS"
End If
End Sub

Public Sub Comando206_Click()
Me.Requery
Me.Filter = "[SI] <> 0 or [INGRESOS] <> 0 or [PROCESADO] <> 0 or [MERMAS] <> 0 or [SF] <> 0"
Me.FilterOn = True
Me.Printer.Orientation = acPRORLandscape
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Procesamiento vs Ingresos.pdf"
DoCmd.OutputTo acOutputForm, "Control Procesamiento / Ingresos", acFormatPDF, archivo, False

DoCmd.Close acForm, "Control Procesamiento / Ingresos"
End Sub

Private Sub Comando52_Click()
Me.Requery
Me.Filter = "[SI] <> 0 or [INGRESOS] <> 0 or [PROCESADO] <> 0 or [MERMAS] <> 0 or [SF] <> 0"
Me.FilterOn = True
End Sub

Private Sub Comando75_Click()
DoCmd.OpenForm "Historial Procesamiento", , , "[CodCotizacion]=" & Me.CodCotizacion & " AND [semana]=" & Me.asemana
End Sub

Private Sub Comando94_Click()
DoCmd.OpenForm "Control Porcionamiento", acNormal, , "[CodIngrediente]='" & Me.CodIngrediente & "'"
[Forms]![Control Porcionamiento]![asemana] = Me.asemana
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.semanaant = numerosemana(Date)
End Sub
