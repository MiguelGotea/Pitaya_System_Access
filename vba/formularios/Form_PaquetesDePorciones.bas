' ==========================================================
' Modulo  : Form_PaquetesDePorciones
' Tipo    : 100
' Lineas  : 335
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando118_Click()
On Error GoTo Nulo
Dim cetiq As Integer
Dim cont As Integer

Me.capacidadimpresion = Me.Capacidad
Me.nombreimpresion = Me.Nombre & " " & Me.Capacidad
Me.cantidadxpaquete = InputBox("Porciones dentro de paquete:", "Ingresar Datos", 10)
cetiq = InputBox("Cantidad de etiquetas", "Ingresar Datos", 1)
cont = 1

Do While cont <= cetiq
    DoCmd.OpenReport "SoloQRPaquete", acViewNormal
    cont = cont + 1
Loop
Exit Sub
Nulo:
End Sub

Private Sub Comando124_Click()
On Error GoTo Nulo
Dim cetiq As Integer

Me.capacidadimpresion = Me.Capacidad
Me.nombreimpresion = Me.Nombre & " " & Me.Capacidad
Me.cantidadxpaquete = Me.PaquetePorciones
cetiq = InputBox("Cantidad de hojas a imprimir", "Ingresar Datos", 1)

DoCmd.OpenReport "EtiquetasPorcionesGrande", acViewReport
DoCmd.PrintOut acPages, 1, 1, , cetiq, True
DoCmd.Close acReport, "EtiquetasPorcionesGrande"
Exit Sub
Nulo:
End Sub

Private Sub Comando14_Click()
On Error GoTo Nulo

Me.capacidadimpresion = Me.Capacidad
Me.nombreimpresion = Me.Nombre & " " & Me.Capacidad
Me.cantidadxpaquete = Me.PaquetePorciones

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)
Exit Sub
Nulo:
End Sub


Private Sub Comando168_Click()
On Error GoTo Nulo

Me.cantidadxpaquete = "Proteina"
Me.solonombreimpresion = "Proteina"
Me.nombreimpresion = ""
Me.capacidadimpresion = "1 libra"
Me.especiales = "ESPECIAL"

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)

Exit Sub
Nulo:
End Sub

Private Sub Comando171_Click()
On Error GoTo Nulo

Me.cantidadxpaquete = "Mani Horneado"
Me.solonombreimpresion = "Mani Horneado"
Me.nombreimpresion = ""
Me.capacidadimpresion = "1 libra"
Me.especiales = "ESPECIAL"

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)

Exit Sub
Nulo:
End Sub

Private Sub Comando59_Click()


Me.cantidadxpaquete = Me.PaquetePorciones
Me.solonombreimpresion = Me.Nombre
Me.capacidadimpresion = Me.Capacidad
Me.conversionimpresion = Me.Conversion

On Error GoTo Nulo

'definiendo impresora
Dim strDefaultPrinter As String

strDefaulPrintert = Application.Printer.DeviceName
Set Application.Printer = Application.Printers("HPRT D21")

Call configurarmargentamañoreporte("StickerPorciones")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de sticker para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "StickerPorciones", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

Set Application.Printer = Application.Printers(strDefaultPrinter)
Exit Sub
Nulo:
End Sub





Private Sub Comando86_Click() 'ROtulo GRANOLA 1 LB
On Error GoTo Nulo

Me.cantidadxpaquete = "Granola"
Me.solonombreimpresion = "Granola"
Me.nombreimpresion = ""
Me.capacidadimpresion = "1 libra"
Me.especiales = "ESPECIAL"

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)

Exit Sub
Nulo:

End Sub

Private Sub Comando87_Click() 'Avena 1 lb
On Error GoTo Nulo

Me.cantidadxpaquete = "Avena"
Me.solonombreimpresion = "Avena"
Me.nombreimpresion = ""
Me.capacidadimpresion = "1 libra"
Me.especiales = "ESPECIAL"

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)

Exit Sub
Nulo:

End Sub

Private Sub Comando93_Click() 'Cocoa en Polvo 6lb
On Error GoTo Nulo

Me.cantidadxpaquete = "Cocoa"
Me.solonombreimpresion = "Cocoa"
Me.nombreimpresion = ""
Me.capacidadimpresion = "6 libras"
Me.especiales = "ESPECIAL"

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)

Exit Sub
Nulo:

End Sub


Private Sub Comando162_Click() ' Semilla de cacao 6lb
On Error GoTo Nulo

Me.cantidadxpaquete = "S de Cacao"
Me.solonombreimpresion = "S de Cacao"
Me.nombreimpresion = ""
Me.capacidadimpresion = "6 libras"
Me.especiales = "ESPECIAL"

'definiendo impresora
'Dim strDefaultPrinter As String

'strDefaulPrintert = Application.Printer.DeviceName
'Set Application.Printer = Application.Printers("EPSON TM-T20II Receipt5")

Call configurarmargentamañoreporte("PaquetesPorcionesSinCodigo")

'Imprimir
Dim csti As Integer
Dim cont As Integer

csti = InputBox("Cantidad de rotulos para esta medida:", "Ingresar Datos", 1)
Me.cantidadxsticker = csti
cont = 1

Do While cont <= csti
    DoCmd.OpenReport "PaquetesPorcionesSinCodigo", acViewNormal 'acViewNormal
    cont = cont + 1
Loop

'Set Application.Printer = Application.Printers(strDefaultPrinter)

Exit Sub
Nulo:
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
