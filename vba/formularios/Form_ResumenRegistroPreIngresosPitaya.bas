' ==========================================================
' Modulo  : Form_ResumenRegistroPreIngresosPitaya
' Tipo    : 100
' Lineas  : 69
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database

Public Sub Comando214_Click()

Me.Filter = "[Fecha]=#" & Me.fechaac & "# AND [Destino]='" & Me.sucursalac & "' And PorcionDentroDeMezcla([CodCotizacion])=0"
Me.FilterOn = True


Call guardarregistropreingreso

End Sub

Public Sub guardarregistropreingreso()

'Carpeta para archivo
Dim Nombre As String

Nombre = "Despacho Diario " & Month(Me.fechaac) & "_" & Day(Me.fechaac) & "_ " & Year(Me.fechaac) & " - " & Me.sucursalac

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Despacho\" & numerosemana(Me.fechaac) & "\Resumen de Despacho"
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Despacho\" & numerosemana(Me.fechaac)
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Nombre & ".pdf"
DoCmd.OutputTo acOutputForm, "ResumenRegistroPreIngresosPitaya", acFormatPDF, archivo, True

'Sleep (5000)

' Obtener el código del sistema
'Dim CodSistema As Long
'CodSistema = Split(Me.sucursalac, " ")(1)

' Obtener el número del operario con cargo líder en la sucursal actual
'Dim Numero As String
'Numero = DLookup("[Celular]", "[Operarios]", "[Sucursal]=" & CodSistema & " AND [Cargo]='Lider'")

' Mensaje para enviar
'Dim Mensaje As String
'Mensaje = "Buenas, envío del resumen de los preingresos"

' Llamar a la función para enviar la imagen por WhatsApp
'Call EnviarArchivoWhatsapp(Mensaje, archivo, Numero)

End Sub

Private Sub Comando493_Click()

Me.Filter = "[Fecha]=#" & Me.fechaac & "# AND [Destino]='" & Me.sucursalac & "' And PorcionDentroDeMezcla([CodCotizacion])=0"
Me.FilterOn = True


End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

